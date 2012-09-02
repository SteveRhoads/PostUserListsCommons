require 'pony'  # for SMTP e-mails
require 'xml'
require 'active_record'

@user_name = "someuser@domain.com"
@password  = "password"
@rally_url = "rally1.rallydev.com"
cfg_file   = "../_cfg/MyVars_PostUserList.rb"

# Load (and maybe override with) my personal/private variables from a file, if the file exists...
if FileTest.exist?( cfg_file )
    print "Loading <#{cfg_file}>...\n"
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

class RallyWSAPI

  require 'net/https'
  require 'cgi'

  def initialize(rally_user, rally_pwd, rally_url = 'rally1.rallydev.com')
    @http = Net::HTTP.new(rally_url, 443)
    @http.use_ssl = true
    @http.verify_mode = OpenSSL::SSL::VERIFY_NONE
    @rally_user = rally_user
    @rally_pwd = rally_pwd
  end

  def get_user_list(query='(Disabled = false)')
    rwapi       = ""
    user_xml    = ""
    doc         = ""
    content     = "" # set scope to make variable available
    users       = "" # set scope to make variable available
    user_count  = 0  # total per Rally per the query
    user_list   = [] # this is what we are going to package up the rows into

    # -----------------------------------------------------------------------------
    # priming read to get count for this query ------------------------------------
    begin
      user_xml  = self.users(query,pagesize = 1,start = 1,fetch = false)
    rescue Exception => e
      puts "Getting XML: #{e.message}"
    end

    begin
      # Parse the XML with LibXml
      # (this is where the magic happens)
      source  = XML::Parser.string(user_xml) # source.class => LibXML::XML::Parser
      content = source.parse # content.class => LibXML::XML::Document
    rescue Exception => e
      puts "Parse user_xml: #{e.message}"
    end

    # get the number of results found
    begin
      user_count = content.root.find_first('./TotalResultCount').content.to_i
    rescue Exception => e
      puts "Error Getting Record Count: #{e.message}"
      user_count = 0
    end

	# Clear out the SQL Server Database table of the RALLY RECORDS before refilling. Rally's application_id == 1
	ActiveRecord::Base.connection.execute("DELETE FROM [Staging].[app_mgt].[application_user] WHERE [application_id] = 1")
	puts "Destination Table[app_mgt].[application_user]: Rally Records Deleted"

    # Echo user_count
    puts "UserCount: #{user_count.to_s}"

    # Calculate the number of pages (all the full page + the final page.)
    # (user_count / pagesize) + 1)
    pagesize = 200
    pages = (user_count / pagesize) + 1
    # now get all the pages
    pages.times do |page|
      # increments from 0
      puts "Working on page #{page.to_s}"
      begin
        user_xml  = self.users(query,pagesize,start = ((page * pagesize) + 1),fetch = true)
      rescue Exception => e
        puts "Getting a page of XML(#{page.to_s}): #{e.message}"
      end

      # make the REXML xml doc ------------------------
      begin
        source  = XML::Parser.string(user_xml) # source.class => LibXML::XML::Parser
        content = source.parse # content.class => LibXML::XML::Document
      rescue Exception => e
        puts "Parsing a page #{page.to_s} of user_xml: #{e.message}"
      end

      # parse the set and add to the array
      begin
        users = content.root.find('//QueryResult/Results/Object') # entries.class => LibXML::XML::XPath::Object

        users.each do |entry| # entry.class => LibXML::XML::Node
          begin
            # extract the data
            user_id       = (entry.find_first('ObjectID').content != nil) ? entry.find_first('ObjectID').content : ""
            display_name  = (entry.find_first('UserName').content != nil) ? entry.find_first('UserName').content : ""
            email_address = (entry.find_first('EmailAddress').content != nil) ? entry.find_first('EmailAddress').content : ""

            begin
              # write to file
              user_list << "#{user_id},#{display_name},#{email_address}"
            rescue Exception => e
              puts "Error writing user list: #{e.message}"
            else
            ensure
            end

            begin
              # Update SQL Server - Reporting Server
              au = ApplicationUser.new
              au.application_id           = 1  # Commons Application_ID
              au.user_id                  = user_id
              au.user_name                = display_name
              au.email                    = email_address
              au.update_date              = Time.now.to_s
              au.save
            rescue Exception => e
              puts "Error updating SQL Server Reporting Server[app_mgt].[application_user]: #{e.message}"
            else
            ensure
            end

            # DEBUG #Advise Completion on Console
            # puts "Record added: #{display_name}"
          rescue Exception => e
              puts "Problem processing XML (to file, console and SQL Server): #{user_id}: #{e.message}"		 			  
          end
        end
      rescue Exception => e
        puts "Reading a page (#{page.to_s}) of user rows: #{e.message}"
      end
    end
    return user_list
  end

  def users(query="", pagesize=20, start=1, fetch=true)

    # https://rally1.rallydev.com/slm/webservice/1.27/user.js?workspace=https://rally1.rallydev.com/slm/webservice/1.27/workspace/343331994&query=&fetch=true&start=1&pagesize=20
    # /slm/webservice/1.27/user.js?workspace=https://rally1.rallydev.com/slm/webservice/1.27/workspace/343331994&query=&fetch=true&start=1&pagesize=20
    #/slm/webservice/1.27/user.js?query=fetch=true&start=1&pagesize=200

    @uri_rest_root            = '/slm/webservice'
    @uri_rest_api_version     = '/1.27'
    @uri_rest_artifact        = '/user'
    @uri_rest_query_parm      = '?query='
    @uri_rest_query_value     = CGI::escape(query.to_s) # CGI::unescape()
    @uri_rest_fetch_parm      = '&fetch='
    @uri_rest_fetch_value     = fetch.to_s
    @uri_rest_pagesize_parm   = '&pagesize='
    @uri_rest_pagesize_value  = pagesize.to_s
    @uri_rest_pagestart_parm  = '&start='
    @uri_rest_pagestart_value = start.to_s

    @uri_rest_get = @uri_rest_root  +
        @uri_rest_api_version      +
        @uri_rest_artifact         +
        @uri_rest_query_parm       +
        @uri_rest_query_value      +
        @uri_rest_fetch_parm       +
        @uri_rest_fetch_value      +
        @uri_rest_pagesize_parm    +
        @uri_rest_pagesize_value   +
        @uri_rest_pagestart_parm   +
        @uri_rest_pagestart_value

    puts "REST GET: #{@uri_rest_get}"

    # we make an HTTP basic auth by passing the username and password
    @http.start do |http|
       req = Net::HTTP::Get.new(@uri_rest_get)
       req.basic_auth @rally_user, @rally_pwd

       resp, data = @http.request(req)
       puts "Response Code: #{resp}"
       data    # the raw XML
    end
  end
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

# +++++++++++++++++++++++++++++
@start_time = Time.now
puts "Starting: #{@start_time.to_s}"

list = []
begin
  rwsapi  = RallyWSAPI.new(@rally_user,@rally_pwd, @rally_url)
  list    = rwsapi.get_user_list(query="(Disabled = false)")
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
    files_attached = attach_commons_file(@commons_server, @deaa_userlist_doc, output_file_name, @commons_user, @commons_pwd)
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