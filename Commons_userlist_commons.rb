require 'pony'  # for SMTP e-mails
require 'active_record' # for direct access to the SQL Server - Reporting Server
require 'logger'

# Logger ------------------------------------------------------------
logfile = File.dirname(__FILE__) + "/./commons_users.log"
@logger 		= Logger.new(logfile,'daily')
@logger.level 	= Logger::INFO #FATAL | DEBUG

def log(severity='INFO',msg="")
  puts "#{Time.now.to_s}- #{severity.upcase.chomp}: #{msg.to_s}"
  case severity.upcase
    when 'INFO'
      @logger.info msg.to_s
    when 'WARN'
      @logger.warn msg.to_s
    when 'FATAL'
      @logger.fatal msg.to_s
    when 'DEBUG'
      @logger.debug msg.to_s
    else
      @logger.info msg.to_s
  end
  return msg
end

cfg_file          = File.dirname(__FILE__) + "/../_cfg/MyVars_PostUserList.rb"

user_count        = 0
user_threshold    = 1000

# user accounts ------------
@ora_user          = ""
@ora_pwd           = ""
@commons_user      = ""
@commons_pwd       = ""

# tnsnames -----------------
@ora_commons       = "JIVE"
@users_file        = "commons_users.csv"
@commons_server    = "commons.cable.comcast.com"
@deaa_userlist_doc = "DOC-10361" # test doc

start_time = Time.now
log "INFO", "Starting : #{start_time.to_s}"

# Load (and maybe override with) my personal/private variables from a file, if the file exists...
if FileTest.exist?( cfg_file )
    log "INFO", "Loading <#{cfg_file}>"
    require cfg_file
end

# for interacting with the SQL Server - Reporting Server
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

def create_commons_user_file(ora_user, ora_pwd, ora_instance, output_file_name)
  begin
    require 'oci8'
    # Create User file
    # remove all records without e-mail addresses
    # If there are many e-mails, only keep the first one
    # remove leading and trailing blank space
    # set the array position of the data elements
    user_id   = 0
    user_name = 1
    email_col = 2

    # open the output file for writing
    commons_output_file = File.new("./#{output_file_name}","w")

    # send to file -----------------------------
    commons_output_file.puts '"UserID","UserName","Email"'

    # Clear out the SQL Server Database table before refilling. Common's application_id == 2
    ActiveRecord::Base.connection.execute("DELETE FROM [Staging].[app_mgt].[application_user] WHERE [application_id] = 2")
    log "INFO", "Destination Table[application_user]: Commons Records Deleted"

    # Initialize my counters and connectors
    row_count = 0
    ora_conn = ""
    begin
      ora_conn = OCI8.new(ora_user,ora_pwd,ora_instance)
    rescue DBI::DatabaseError => e
      log "ERROR", "DBI Error: #{e.message}"
    rescue Exception => e
      log "ERROR", "General Error connecting: #{e.message}"
    else
    ensure
      log "INFO", "Connection: #{ora_conn}"
    end

    # Set up resultset sql -----------------------------------------------------
    user_sql  = "SELECT USERID, USERNAME, EMAIL "
    user_sql += "FROM JIVEUSER.JIVEUSER "
    user_sql += "WHERE USERENABLED = 1 AND LASTLOGGEDIN <> 0"

    # Fetch results ------------------------------------------------------------
    ora_results = ora_conn.exec(user_sql)
    ora_results.fetch do |row|
      row_count += 1
      if row[email_col] =~ /(,|;)/       # if embedded comma or semi-colon, may be a list
        email = row[email_col][0..(row[email_col].index(/(,|;)/)-1)].downcase.chomp.gsub(/[^a-z0-9\.@_\-]+/i,'') # only take first email
      else
        email = row[email_col].downcase.chomp.gsub(/[^a-z0-9\.@_\-]+/i,'') # ensure that all email are lower-case.
      end

      # sanitize -------------------------------
      user_id_str = row[user_id].to_s.gsub(/[^a-z0-9\.@_\-]+/i,'').chomp
      user_name_str = row[user_name].to_s.downcase.gsub(/[^a-z0-9\.@_\-]+/i,'').chomp

      # show on console -------------------------------------------------------
      log "INFO", '"'                      +
         user_id_str + '","'        +
         user_name_str + '","'      +
         email                      + '"'

      # send to file -----------------------------
      commons_output_file.puts '"'  +
         user_id_str + '","'        +
         user_name_str + '","'      +
         email                      + '"'

      # Update SQL Server - Reporting Server
      au = ApplicationUser.new
        au.application_id           = 2  # Commons Application_ID
        au.user_id                  = user_id_str
        au.user_name                = user_name_str
        au.email                    = email
        au.update_date              = Time.now.to_s
      au.save

      # Advise Completion on Console
      log "INFO", "Record added: #{user_name_str}"
    end

    # Close the output file
    commons_output_file.close
  rescue Exception => e
	log "ERROR", "General error: #{e.message}"
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

  log "INFO", "Attaching Commons file - response code: #{response.code}"

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
    log "INFO", "Attachment Names: #{ele.elements["name"]}"
    if ele.elements["name"].text.to_s == attachment_file_name + ".zip"
        log "INFO", "Deleting existing attachment #"
        del   =  Net::HTTP::Delete.new("/rpc/rest/documentService/attachments/#{ele.elements["ID"].text}")
        del.basic_auth(commons_user, commons_pwd)		# authentication to Commons
        response = http.request(del)
        response.code
    end
  end

  # now add the attachment
  log "INFO", "Adding fresh attachment: ./#{attachment_file_name}"
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
log "INFO", "Oracle Login: #{@ora_user_commons}"

begin
  user_count = create_commons_user_file(@ora_user_commons, @ora_pwd_commons, @ora_commons_prod, @commons_users_file)
rescue Exception => e
  log "ERROR", "Error creating user file and updating ReportingServer [Staging].[application_user]: #{e.message}"
  user_count = 0
else
ensure
  log "INFO", "#{user_count} users written"
end

# if it looks like we connected to Oracle and got our list, go ahead and post the results
begin
  if user_count > user_threshold
    files_attached = attach_commons_file(@commons_server, @deaa_userlist_doc, @commons_users_file, @commons_user, @commons_pwd)
    log "INFO" "#{files_attached} Files attached"
  else
	log "WARN", "User Count(#{user_count.to_s}) below threshold. Did NOT attach new file"
  end
rescue Exception => e
  log "ERROR", "Error attaching file to Commons: #{e.message}"
else
ensure
end

# Update application_user_daily for Commons Users
app_id    = 2
data_date = Time.now.strftime("%Y-%m-%d")
begin
  au_sql = "EXEC [Staging].[app_mgt].[create_or_update_app_user_daily] @application_id=#{app_id.to_s},@data_date='#{data_date.to_s}',@user_count=#{user_count.to_s}"
  ActiveRecord::Base.connection.execute(au_sql)
  log "INFO", "Updated Daily Commons User Count(#{user_count.to_s})"
rescue Exception => e
  log "ERROR", "Error updating Daily Commons User Count: #{e.message}"
else
ensure
end

# Update production with the newest values
begin
  au_sql = "EXEC [Staging].[app_mgt].[application_user_daily_move_staging_to_prod]"
  ActiveRecord::Base.connection.execute(au_sql)
  log "INFO", "Updated Production: Daily App User Count"
rescue Exception => e
  log "INFO", "Error updating Production: Daily App User Count: #{e.message}"
else
ensure
end

# send completion e-mail
# E-mail options -----------------------------------------------
watchers    = ["steve_rhoads@cable.comcast.com",
               "steve.rhoads@radcorp.com"]

cc          = ["Jim<james_wentz@cable.comcast.com>",
               "Shalene<shalene_copeland@cable.comcast.com>",
               "Peter<peter_hart@cable.comcast.com>"]

cc_list     = cc.join(",")

body_str = "Commons User List uploaded to Commons(#{user_count.to_s})."

# prepare and send the email
email_subject = ""
if user_count > user_threshold then
  email_subject = "Commons List(#{user_count.to_s}) Successfully Uploaded #{Time.now.to_s}"
else
  email_subject = "Commons List(#{user_count.to_s}) Failed to Upload #{Time.now.to_s}"
end

Pony.options = {  :from         => 'Steve Rhoads <steve_rhoads@cable.comcast.com>',
				  :via          => :smtp,
                  #:cc           => cc_list,
                  :subject      => email_subject,
                  :body         => body_str,
				  :via_options  => { :address => 'apprelay.css.cable.comcast.com' } }

watchers.each do |email_to|
  begin
    Pony.mail(:to => email_to)
	log "INFO", "Mail Sent! (#{Time.now.to_s})"
  rescue Exception => e
    log "ERROR", "Problem during mailing(#{email_to}) - #{e.message}"
  end
end

# Capture our run-time
log "INFO", "Elapsed #{'%.1f' % ((Time.now - start_time)/60)} min"