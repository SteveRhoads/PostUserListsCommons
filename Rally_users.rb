require 'pony'  # for SMTP e-mails
require 'xml'
require 'active_record'
require 'rally_api'

# capture start time
@start_time = Time.now

# Load (and maybe override with) my personal/private variables from a file, if the file exists...
cfg_file   = File.dirname(__FILE__) + "/../_cfg/MyVars_PostUserList.rb"
if FileTest.exist?( cfg_file )
  puts "Loading <#{cfg_file}>..."
  require cfg_file
end

puts "Parameters: #{cfg_file},#{@base_url},#{@rally_user},#{@rally_pwd}"

# for interacting with the SQL Server - Reporting Server
class ApplicationUser < ActiveRecord::Base
  self.table_name = "app_mgt.application_user"
end

ActiveRecord::Base.establish_connection(
    :adapter  =>  "sqlserver",
    :host     =>  @reporting_server_host,
    :database =>  @reporting_server_db,
    :username =>  @reporting_server_account,
    :password =>  @reporting_server_pwd
)

class RallyJson

  def initialize(rally_user, rally_pwd, base_url = 'https://rally1.rallydev.com/slm')
    # Connect to Rally via the new RallyAPI (json)
    # Setup the Custom Headers
    headers = RallyAPI::CustomHttpHeader.new()
    headers.name    = "Rally User List"
    headers.vendor  = "Comcast PE-SED-DEAA"
    headers.version = "1.0.1"

    # Setup connection configuration
    config = {:base_url => base_url,
              :username => rally_user,
              :password => rally_pwd,
              :headers  => headers}

    # ESTABLISH connection, using the configuration
    @rally  = RallyAPI::RallyRestJson.new(config)
  end

  def get_user_list(query='(Disabled = false)')
    rwapi       = ""
    user_xml    = ""
    doc         = ""
    content     = "" # set scope to make variable available
    users       = "" # set scope to make variable available
    user_count  = 0  # total per Rally per the query
    user_list   = [] # this is what we are going to package up the rows into

    # Setup QUERY
    user_query   = RallyAPI::RallyQuery.new()
    user_query.type          = :user
    user_query.fetch         = "ObjectID,UserName,EmailAddress"
    user_query.order         = "UserName Asc"
    user_query.query_string  = query

    # EXECUTE query
    results = @rally.find(user_query)

    # Clear out the SQL Server Database table of the RALLY RECORDS before refilling. Rally's application_id == 1
    ActiveRecord::Base.connection.execute("DELETE FROM [Staging].[app_mgt].[application_user] WHERE [application_id] = 1")
    puts "Destination Table[app_mgt].[application_user]: Rally Records Deleted"

    # Echo user_count
    # puts "UserCount: #{results.total_result_count.to_s}"

    # Cycle through Rally User - Results List
    results.each do |r|
      begin
        # write to file
        user_list << "#{r.ObjectID},#{r.DisplayName.to_s},#{r.EmailAddress}"

        # Update Staging - SQL Server
        begin
          au = ApplicationUser.new
          au.application_id           = 1  # Commons Application_ID
          au.user_id                  = r.ObjectID
          au.user_name                = r.UserName.to_s
          au.email                    = r.EmailAddress
          au.update_date              = Time.now.to_s
          au.save
        rescue Exception => e
          puts "Error updating SQL Server Reporting Server[app_mgt].[application_user]: #{e.message}"
        else
        ensure
        end

      rescue Exception => e
        puts "Error writing user list: #{e.message}"
      else
      ensure
      end
    end
    # return the list of user accounts (ObjectID,DisplayName,EmailAddress)
    return user_list
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
end

# +++++++++++++++++++++++++++++
@start_time = Time.now
puts "Starting: #{@start_time.to_s}"

list = []
rj = ""
begin
  rj    = RallyJson.new(@rally_user, @rally_pwd, @base_url)
  list  = rj.get_user_list(query="(Disabled = false)")
rescue Exception => e
  puts "Connecting to Rally: #{e.message}"
end

# open the output file for writing
output_file_name = "rally_users.csv"
output_file = File.new("./#{output_file_name}","w")

user_count = list.length

puts "Users:#{user_count.to_s} "
list.each do |l|
  #puts l
  begin
    output_file.puts l
  rescue Exception => e
    puts "Error writing file: #{e.message}"
  else
  ensure
  end
end
output_file.close

# if it looks like we connected to Oracle and got our list, go ahead and post the results
user_threshold = 1000 # if we got 1000 users, we had a good connection
begin
  if user_count > user_threshold
    files_attached = rj.attach_commons_file(@commons_server, @deaa_userlist_doc, output_file_name, @commons_user, @commons_pwd)
    puts "#{files_attached} Files attached"
  end
rescue Exception => e
  puts "Error attaching file to Commons: #{e.message}"
else
ensure
end

# Update [Production] from [Staging].[app_mgt].[application_user]
# send to SQL Server
begin
  au_sql = "EXEC [Staging].[app_mgt].[application_user_move_staging_to_prod]"
  ActiveRecord::Base.connection.execute(au_sql)
  puts "Moved [staging].[app_mgt].[application_user] to [production]"
rescue Exception => e
  puts "Error updating [Production] from [Staging].[app_mgt].[application_user]: #{e.message}"
else
ensure
end

# Update application_user_daily for Rally Users
app_id       = 1
data_date    = Time.now.strftime("%Y-%m-%d")
begin
  au_sql = "EXEC [Staging].[app_mgt].[create_or_update_app_user_daily] @application_id=#{app_id.to_s},@data_date='#{data_date.to_s}',@user_count=#{user_count.to_s}"
  ActiveRecord::Base.connection.execute(au_sql)
  puts "Updated Daily Rally User Count"
rescue Exception => e
  puts "Error updating Daily Rally User Count: #{e.message}"
else
ensure
end

# Update production with the newest values
begin
  au_sql = "EXEC [Staging].[app_mgt].[application_user_daily_move_staging_to_prod]"
  ActiveRecord::Base.connection.execute(au_sql)
  puts "Updated Production: Daily App User Count"
rescue Exception => e
  puts "Error updating Production: Daily App User Count: #{e.message}"
else
ensure
end

# send completion e-mail
# E-mail options -----------------------------------------------
email_subject = ""
if user_count > user_threshold then
  email_subject = "Rally User List uploaded to Commons #{Time.now.to_s} Users=#{user_count}"
else
  email_subject = "Rally User List Failed to Upload to Commons #{Time.now.to_s} Users=#{user_count}"
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
                  #:cc           => cc_list,
                  :subject      => email_subject,
                  :body         => body_str,
				          :via_options  => { :address => 'apprelay.css.cable.comcast.com' } }

watchers.each do |email_to|
  begin
    Pony.mail(:to => email_to)
  rescue Exception => e
    "Problem during mailing(#{email_to}) - #{e.message}"
  end
end

# End of script ------------------------------------------------
puts "Elapsed: #{'%.1f' % ((Time.now - @start_time)/60)} minutes"