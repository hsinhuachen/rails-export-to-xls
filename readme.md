# Rails export to excel

## link_to export
	# view
	link_to "Export", projects_path(format: :xls, params: request.parameters), class: "btn btn-default"

## Controller
	def index
		@projects = Project.all
			
		respond_to do |format|
		 	format.html
		 	format.xls
		 	format.csv { send_data @projects.to_csv, filename: "project-#{Date.today}.csv" }
		 end
	end

## Config
	# config/initializers/mime_types.rb
	Mime::Type.register "application/xls", :xls


## XML format
	# view/projects/index.xls.html
	<?xml version='1.0'?>
	<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
	  xmlns:o="urn:schemas-microsoft-com:office:office"
	  xmlns:x="urn:schemas-microsoft-com:office:excel"
	  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
	  xmlns:html="http://www.w3.org/TR/REC-html40">
	  <Worksheet ss:Name="Projects">
	    <Table>
	      <Row>
	        <Cell><Data ss:Type="String">編號</Data></Cell>
	        <Cell><Data ss:Type="String">標題</Data></Cell>
	      </Row>
	      <% @projects.each do |project| %>
	        <Row>
	          <Cell><Data ss:Type="String"><%= project.id %></Data></Cell>
	          <Cell><Data ss:Type="String"><%= project.title %></Data></Cell>
	        </Row>
	      <% end %>
	    </Table>
	  </Worksheet>
	</Workbook>
