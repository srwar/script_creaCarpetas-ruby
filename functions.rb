def name_month(month_number)
	months = {
		1 => 'ENE',
		2 => 'FEB',
		3 => 'MAR',
		4 => 'ABR',
		5 => 'MAY',
		6 => 'JUN',
		7 => 'JUL',
		8 => 'AGO',
		9 => 'SEP',
		10 => 'OCT',
		11 => 'NOV',
		12 => 'DIC'
	}
	if months.has_key?(month_number)
		return months[month_number]
	end
	if month_number == 0
		return months[12]
	end
end