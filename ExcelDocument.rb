require 'win32ole'

class ExcelDocument
  def initialize
    @excel = WIN32OLE::new('excel.Application')
    @line = 0
    @startRow = 0
  end
  
  def openFile(filePath)
    @workbook = @excel.Workbooks.open(filePath)
    @worksheet = @workbook.worksheets(1)
  end
  
  def switchWorksheet(sheetName)
    @worksheet = @workbook.worksheets(sheetName)
  end
  
  def nextRow()
    @line = @line + 1
    return @worksheet.Range("A#{@line}").Value
  end
  
  def topRow()
    @line = @startRow
    return @worksheet.Range("A#{@line}").Value
  end
  
  def setStartRow(line)
    @startRow = line
    @line = line
  end
  
  def getCurrentRow()
    return @line
  end
  
  def getValue(colName, rowNum)
    return @worksheet.Range("#{colName}#{rowNum}").Value
  end
  
  def closeWorkbook()
    @workbook.close
  end
end
