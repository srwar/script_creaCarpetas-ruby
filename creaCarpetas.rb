require_relative "functions"
require "progress_bar"
require "spreadsheet"
Spreadsheet.client_encoding = 'UTF-8'

#Open file
folders = File.read("carpetas.txt").split
#Open excel 
book = Spreadsheet.open 'DEBITOS.xls'
#Select the first worksheet of the workbook
sheet1 = book.worksheet 0
#Select Month and Year for the new book
time = Time.new
#What gonna happen when month = 1?
month = name_month(time.month-1)
year = time.year
puts "Processing: #{month}/#{year}"
puts "Starting..."
#New progress bar
bar = ProgressBar.new(folders.count, :bar, :counter, :percentage, :elapsed, :rate)
cuie=''
row = sheet1.row(2)
#Set cell's format (red and yellow)
format = Spreadsheet::Format.new :color=> :red, :pattern_fg_color => :yellow, :pattern => 2
coincidences=0
#starting = Process.clock_gettime(Process::CLOCK_MONOTONIC)
#Create folders
folders.each do |folder|
	#Create a new Excel 
	new_book = Spreadsheet::Workbook.new
	#Create two sheets
	sheet = new_book.create_worksheet(name: 'Prestaciones')
	sheet_debitos = new_book.create_worksheet(name: 'DEBITOS')
	#Add headers
	sheet.row(0).push('id_comprobante', 'fecha_comprobante', 'periodo', 'clavebeneficiario', 'ApellidoBenef', 'NombreBenef', 'FechaNacimientoBenef', 'NroDocumentoBenef', 'Edad', 'Mes', 'SexoBeneficiario', 'cantidad', 'codigo', 'diagnostico', 'embarazo_actual', 'semanas_embarazo', 'fecha_diagnostico', 'fecha_probable_de_parto', 'VerificaSiTienePrestacionesAsociadas_descripcion', 'EquivalenciasNomenActu_descripcion', 'altacomp', 'precio_prestacion', 'precio', 'id_nomenclador', 'id_nomenclador_nuevo', 'usuario', 'id_prestacion', 'cuie', 'id_factura', 'Cod_Error', 'VerificaAntesdeVincularActivos_Descripcion', 'DELETE_id_nomenclador_nuevo')
	sheet_debitos.row(0).push('id_comprobante', 'fecha_comprobante', 'periodo', 'clavebeneficiario', 'ApellidoBenef', 'NombreBenef', 'FechaNacimientoBenef', 'NroDocumentoBenef', 'Edad', 'Mes', 'SexoBeneficiario', 'cantidad', 'codigo', 'diagnostico', 'embarazo_actual', 'semanas_embarazo', 'fecha_diagnostico', 'fecha_probable_de_parto', 'VerificaSiTienePrestacionesAsociadas_descripcion', 'EquivalenciasNomenActu_descripcion', 'altacomp', 'precio_prestacion', 'precio', 'id_nomenclador', 'id_nomenclador_nuevo', 'usuario', 'id_prestacion', 'cuie', 'id_factura', 'Cod_Error', 'VerificaAntesdeVincularActivos_Descripcion', 'DELETE_id_nomenclador_nuevo', 'motivo de baja')
	#Create name of folder
	factura = "factura" + folder
	#Create directory
	Dir.mkdir(factura)
	i=0
	sheet1.each_with_index 2 do |row|
		if row[28].to_i == folder.to_i
			cuie = row[27]
			#Insert all fields in the new book
			sheet.insert_row i+1, [ row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[11],row[12],row[13],row[14],row[15],row[16],row[17],row[18],row[19],row[20],row[21],row[22],row[23],row[24],row[25],row[26],row[27],row[28],row[29],row[30],row[31] ]
			#color and copy duplicate cells
			if row[32] == 1
				sheet.row(1).default_format = format #for entire first row
				sheet_debitos.insert_row i+1, [ row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[11],row[12],row[13],row[14],row[15],row[16],row[17],row[18],row[19],row[20],row[21],row[22],row[23],row[24],row[25],row[26],row[27],row[28],row[29],row[30],row[31] ]
				sheet_debitos.row(1).default_format = format #for entire first row
			end
			coincidences+=1
		end
	end
	#Save new book
	new_book.write "#{factura}/DEBITOS_#{month}#{year}-#{cuie}.xls"
	system "cls"
	puts "--------------------------------------------------------------------------------"
	puts "                                              count   progress elapsed    rate"
	puts "--------------------------------------------------------------------------------"
	bar.increment!
#puts cuie
#puts coincidences
coincidences=0
end
#ending = Process.clock_gettime(Process::CLOCK_MONOTONIC)
#elapsed = ending - starting
#puts "---------------------------------------------------------"
#puts "#{elapsed} seconds"


