require 'pony'  # for SMTP e-mails
require 'active_record'

cfg_file          = "../_cfg/MyVars_PostUserList.rb"
user_count        = 0
user_threshold    = 5000

# user accounts ------------
@ora_user          = ""
@ora_pwd           = ""
@commons_user      = ""
@commons_pwd       = ""

# tnsnames -----------------
@ora_hpqc          = "qcprddr"
@hpqc_users_file   = "hpqc_users.csv"
@commons_server    = "commons.cable.comcast.com"
@deaa_userlist_doc = "DOC-10361" # test doc

start_time = Time.now
puts "Starting : #{start_time.to_s}"

# Load (and maybe override with) my personal/private variables from a file, if the file exists...
if FileTest.exist?( cfg_file )
    print "Loading <#{cfg_file}>...\n"
    require cfg_file
end

class ApplicationUser < ActiveRecord::Base
  set_table_name "app_mgt.application_user"
end

ActiveRecord::Base.establish_connection(
    :adapter  =>  "sqlserver",
    :host     =>  @reporting_server_host,
    :database =>  @reporting_server_db,
    :username =>  @reporting_server_account,
    :password =>  @reporting_server_pwd
)

puts "Oracle Login: #{@ora_user}"

def create_hpqc_user_file(ora_user, ora_pwd, ora_instance, output_file_name)
  begin
    require 'oci8'
    # Create HPQC User file
    # remove all records without e-mail addresses
    # If there are many e-mails, only keep the first one
    # remove leading and trailing blank space
    # set the array position of the data elements
    user_id         = 0
    user_name       = 1
    email_col       = 2
    active          = 3
    full_name       = 4
    description     = 5
    phone_number    = 6
    last_update     = 7
    us_dom_auth     = 8
    us_report_role  = 9

    # open the output file for writing
    hpqc_user_file = File.new("./#{output_file_name}","w")

    # Set Headers -----------------------------
    hpqc_user_file.puts '"UserID","UserName","Email"'

    row_count = 0
    # connect to the HPQC Reporting Server (Oracle) to get the user names
    user_sql = "SELECT u.user_id,u.user_name,u.email,u.acc_is_active,u.full_name,u.description,"
    user_sql += "u.phone_number,u.last_update,u.us_dom_auth,u.us_report_role FROM QCSITEADMIN_DB.USERS u "
    user_sql += "WHERE u.EMAIL IS NOT NULL AND u.EMAIL NOT LIKE '%works.com'"

    puts "UserSQL: #{user_sql}"
    puts "Credentials: #{ora_user}, #{ora_instance}"
    ora_conn = ""
    begin
      ora_conn = OCI8.new(ora_user,ora_pwd,ora_instance)
    rescue DBI::DatabaseError => e
      puts "An error occurred"
      puts "Error code:    #{e.err}"
      puts "Error message: #{e.errstr}"
    rescue Exception => e
      "Error connecting: #{e.message}"
    else
    ensure
      puts "Connection: #{ora_conn}"
    end

    # Clear out the SQL Server Database table before refilling
    ActiveRecord::Base.connection.execute("DELETE FROM [Staging].[app_mgt].[application_user] WHERE [application_id] = 3")
    puts "Destination Table[app_mgt].[application_user]: HPQC Records Deleted"

    # Fetch results
    ora_results = nil
    ora_results = ora_conn.exec(user_sql)
    ora_results.fetch do |row|
      row_count += 1
      if row[email_col] =~ /(,|;)/       # if embedded comma or semi-colon, may be a list
        email = row[email_col][0..(row[email_col].index(/(,|;)/)-1)].downcase.gsub(/[^a-z0-9\.@_\-]+/i,'').chomp # only take first email
      else
        email = row[email_col].downcase.gsub(/[^a-z0-9\.@_\-]+/i,'').chomp # ensure that all email are lower-case.
      end

      user_id_str       = row[user_id].to_s.gsub(/[^a-z0-9\.@_\-]+/i,'').chomp.downcase # sanitize!!!
      user_name_str     = row[user_name].to_s.gsub(/[^a-z0-9\.@_\-]+/i,'').chomp.downcase # sanitize!!!
      active_str        = row[active].to_s.gsub(/[^a-z0-9\.@_\-]+/i,'').chomp.downcase # sanitize!!!
      full_name_str     = row[full_name].to_s.gsub(/[^a-z0-9\.@_\-]+/i,'').chomp.downcase # sanitize!!!
      phone_number_str  = row[phone_number].to_s.gsub(/[^a-z0-9\.@_\-]+/i,'').chomp.downcase # sanitize!!!
      last_update_str   = row[last_update].to_s.gsub(/[^a-z0-9\.@_\-]+/i,'').chomp.downcase # sanitize!!!

      # show on console --------------------------
      puts '"'                     +
         user_id_str + '","'       +
         user_name_str + '","'     +
         email  + '"'

      # send to file -----------------------------
      hpqc_user_file.puts '"'      +
         user_id_str + '","'       +
         user_name_str + '","'     +
         email + '"'

      # send to SQL Server
      #au_sql = "INSERT INTO [dev].[app_mgt].[application_user]([application_id],[user_id],[user_name],[email],[update_date]) VALUES(3,'"
      #au_sql += user_id_str + "','" + user_name_str + "','" + email + "','" + Time.now.to_s + "')"
      #puts "INSERT SQL: #{au_sql}"
      #ActiveRecord::Base.connection.execute(au_sql)
      au = ApplicationUser.new
        au.application_id = 3
        au.user_id        = user_id_str
        au.user_name      = user_name_str
        au.email          = email
        au.update_date    = Time.now.to_s
      au.save

      puts "Record added: #{user_name_str}"
    end

   # Close the output file
    hpqc_user_file.close
  rescue Exception => e
    puts "Error retrieving results from Oracle: #{e.message}"
  else
  ensure
  end
  return row_count
end

def attach_commons_file(commons_server, owning_document, attachment_file_name, commons_user, commons_pwd)
# getAttachmentID that matches attachment_name for HPQC_Users.csv
# deleteAttachment_by_ID
# add new attachment

  require 'net/http'
  require 'rexml/document'
  require 'base64'	# needed to Base64 encode the binary attachment

  http  = Net::HTTP.start(commons_server)
  req   = Net::HTTP::Get.new("/rpc/rest/documentService/attachments/#{owning_document}")
  req.basic_auth(commons_user,  commons_pwd)		# authentication to Commons
  req["Content-Type"] = "text/xml"

  # XML response contains the attachment id in "<ID>"
  response = http.request(req)

  puts "Response code: #{response.code}"

  #puts "Response body:#{response.body}"
  # Layout of response ~~~~~~~~~~~~~~~~~~~~~
  # extract event information
  # Elements
  #    return
  #       ID
  #       objectType
  #       uuid
  #       version
  #       contentType
  #       data
  #       name

  doc = REXML::Document.new(response.body)
  doc.elements.each('ns2:getAttachmentsByDocumentIDResponse/return') do |ele|
    puts "Attachment Names: #{ele.elements["name"]}"
    if ele.elements["name"].text.to_s == attachment_file_name + ".zip"
        puts "Deleting attachment #"
        del   =  Net::HTTP::Delete.new("/rpc/rest/documentService/attachments/#{ele.elements["ID"].text}")
        del.basic_auth(commons_user, commons_pwd)		# authentication to Commons
        response = http.request(del)
        response.code
    end
  end

  # now add the attachment
  puts "Adding attachment: ./#{attachment_file_name}"
  req   = Net::HTTP::Post.new('/rpc/rest/documentService/attachments')
  req.basic_auth(commons_user,commons_pwd)		# authentication to Commons

  # grab the file and attach it to the master document
  file = File.open("./#{attachment_file_name}","rb")	# input file
  file_contents = Base64.encode64(file.read)		# read and encode the file so that it can be part of a text message

  xml_request  = nil
  xml_request  =  "<addAttachmentToDocumentByDocumentId><documentID>#{owning_document}</documentID><name>#{attachment_file_name}</name><contentType>text/plain</contentType><source>" + file_contents + "</source></addAttachmentToDocumentByDocumentId>"
  req["Content-Type"] = "text/xml"

  # make the attachment request
  response  = http.request(req, xml_request)
end

# MAIN PROGRAM ******************************************************
# If we got a file -> Continue, otherwise retry? and eventually quit
user_count = create_hpqc_user_file(@ora_user, @ora_pwd, @ora_hpqc, @hpqc_users_file)
puts "#{user_count} users written"


# Connect to Commons and Attach the file
begin
  # if it looks like we connected to Oracle and got our list, go ahead and post the results
  if user_count > user_threshold # the threshold is just a final check that we really got a good update file
    files_attached = attach_commons_file(@commons_server, @deaa_userlist_doc, @hpqc_users_file, @commons_user, @commons_pwd)
    puts "#{files_attached} Files attached"
  end
rescue Exception => e
  puts "Error attaching HPQC userlist to Commons: #{e.message}"
else
ensure
end

# Update [Production] from [Staging].[application_user]
# send to SQL Server
begin
  au_sql = "EXEC [Staging].[app_mgt].[application_user_move_staging_to_prod]"
  ActiveRecord::Base.connection.execute(au_sql)
  puts "Moved [staging].[app_mgt].[application_user] to [production]"
rescue Exception => e
  puts "Error updating [Production] from [Staging].[application_user]: #{e.message}"
else
ensure
end

# Update application_user_daily for HPQC Users
app_id      = 3
data_date   = Time.now.strftime("%Y-%m-%d")
begin
  au_sql = "EXEC [Staging].[app_mgt].[create_or_update_app_user_daily] @application_id=#{app_id.to_s},@data_date='#{data_date.to_s}',@user_count=#{user_count.to_s}"
  ActiveRecord::Base.connection.execute(au_sql)
  puts "Updated Daily HPQC User Count"
rescue Exception => e
  puts "Error updating Daily HPQC User Count: #{e.message}"
else
ensure
end

# Update production with the newest values
begin
  au_sql = "EXEC [Staging].[app_mgt].[application_user_daily_move_staging_to_prod]"
  ActiveRecord::Base.connection.execute(au_sql)
  puts "Updated Production: Daily App User Count"
rescue Exception => e
  puts "Error updating Production: [application_user_daily] User Count: #{e.message}"
else
ensure
end

# send completion e-mail
# E-mail options -----------------------------------------------
email_subject = ""
if user_count > user_threshold then
  email_subject = "HPQC User List Successfully Uploaded to Commons.Time: #{Time.now.to_s} Users: #{user_count.to_s}"
end
watchers    = ["steve_rhoads@cable.comcast.com",
               "steve.rhoads@radcorp.com"]

cc          = ["Jim<james_wentz@cable.comcast.com>",
               "Shalene<shalene_copeland@cable.comcast.com>",
               "Peter<peter_hart@cable.comcast.com>"]

cc_list     = cc.join(",")

body_str = email_subject

# prepare and send the email
Pony.options = {  :from         => 'Steve Rhoads <steve_rhoads@cable.comcast.com>',
				   :via          => :smtp,
                  # :cc           => cc_list,
                  :subject      => email_subject,
                  :body         => body_str,
				          :via_options  => { :address => 'apprelay.css.cable.comcast.com' } }

watchers.each do |email_to|
  begin
    Pony.mail(:to => email_to)
    puts "Mailing to #{email_to}"
  rescue Exception => e
    "Problem during mailing(#{email_to}) - #{e.message}"
  end
end

puts "Elapsed #{'%.1f' % ((Time.now - start_time)/60)} min"
