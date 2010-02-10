class Array
  
  def to_xls(options = {})
    output = '<?xml version="1.0" encoding="UTF-8"?><Workbook xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40" xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office">'
    output << '<Styles>'
			output << '<Style ss:ID="Default" ss:Name="Normal">'
				output << '<Alignment ss:Vertical="Bottom" />'
				output << '<Borders />'
				output << '<Font ss:FontName="Arial" />'
				output << '<Interior />'
				output << '<NumberFormat />'
				output << '<Protection />'
			output << '</Style>'
			output << '<Style ss:ID="s21">'
				output << '<Font x:Family="Swiss" ss:Bold="1" />'
			output << '</Style>'

			output << '<Style ss:ID="s22">'
				output << '<Font ss:Format="General Date" />'
			output << '</Style>'
		output << '</Styles>'
		output << '<Worksheet ss:Name="Sheet1"><Table>'
		
    if self.any?
    
      attributes = self.first.attributes.keys.map { |c| c.to_sym }
      
      if options[:only]
        # the "& attributes" get rid of invalid columns
        columns = options[:only].to_a & attributes
      else
        columns = attributes - options[:except].to_a
      end
    
      columns += options[:methods].to_a
    
      if columns.any?
        unless options[:headers] == false
          output << "<Row>"
          columns.each { |column| output << "<Cell ss:StyleID=\"s21\"><Data ss:Type=\"String\">#{column}</Data></Cell>" }
          output << "</Row>"
        end    

        self.each do |item|
          output << "<Row>"
          columns.each { |column|
						if item.send(column).kind_of?(Integer)
							output << "<Cell><Data ss:Type=\"Number\">#{item.send(column)}</Data></Cell>"
						elsif item.send(column).kind_of?(DateTime)
							output << "<Cell ss:StyleID=\"s22\"><Data ss:Type=\"DateTime\">#{item.send(column).strftime("%Y-%m-%dT%H:%M:%S")}</Data></Cell>"
						else
							output << "<Cell><Data ss:Type=\"String\">#{item.send(column)}</Data></Cell>"
						end
					}
          output << "</Row>"
        end
      end
      
    end
    
    output << '</Table></Worksheet></Workbook>'
  end
  
end
