require 'win32ole'

configMap = Hash.new
secConMap = Hash.new
sectionName = ""

cfgFile = File.new("Configuration.cnfg")

cfgFile.each_line() do |line|
  #line = line.gsub(/\s+/, "")
  line = line.delete(' ')
  if(line !~ /#+(\w*)/ and line !~ /;+(\w*)/ and line !~ /^$/)
    
    if( line =~ /\[(\w+)\]/)
      sectionName = "#{$&}"
      sectionName = sectionName.upcase
      puts "Importing #{sectionName} configuration..."
    elsif line =~ /(\w+)=(\S+)/
      re = /(\w+)=(\w+)/
      md = re.match("#{line}")
      if sectionName =~ /GENERAL/
        configMap["#{md[1]}".upcase]="#{md[2]}".upcase;
        puts configMap["#{md[1]}".upcase]
      elsif sectionName =~ /SECURITYCONTRIBUTION/
        secConMap["#{md[1]}".upcase]="#{md[2]}".upcase;
      end
    end
  end
  
end

cfgFile.close

puts configMap["ImportFile".upcase]

#puts "Please input the path of security position excel file: "
#excel = WIN32OLE::new('excel.Application')
#workbook = excel.Workbooks.open(configMap["ImportFile".upcase])
#worksheet = workbook.worksheets(1)

#line = 2
#while worksheet.Range("a#{line}").Value
#   Portfolio = worksheet.Range("a#{line}").Value
#   puts "Portfolio: "+Portfolio
#   line = line + 1
#end

#workbook.close
