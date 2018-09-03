class ApplicationController < ActionController::Base
	protect_from_forgery with: :exception

	def open_browser
		@browser.close if @browser
		@browser = Watir::Browser.new :chrome
	end

	def open_edge
		@browser.close if @browser
		@browser = Watir::Browser.new :edge
	end

	def open_firefox #http://watir.com/guides/firefox/
		@browser.close if @browser
		@browser = Watir::Browser.new :firefox
	end

	def open_chrome	#https://peter.sh/experiments/chromium-command-line-switches/
		@browser.close if @browser
		prefs = { profile: { managed_default_content_settings: {images: 1} } }
		args = ['--ignore-certificate-errors', '--disable-popup-blocking', '--disable-translate'] #, '--net-log-capture-mode' # --user-data-dir='C:\Users\Tzuyu\AppData\Local\Google\Chrome\User Data']
		@browser = Watir::Browser.new :chrome, options: {prefs: prefs, args: args }
		# @browser.window.resize_to(1000, 800)
	end

	def application_chipset2socket(chipset)
		case chipset
		when "X299"
			return "2066"
		when "X99", "X79"
			return "2011"
		when "Z370", "H370", "B360", "H310", "Z270", "Q270", "H270", "B250", "Z170", "Q170", "H170", "B150", "H110", "C232", "C606"
			return "1151"
		when "Z97", "H97", "Z87", "H87", "B85", "Q87", "H81", "C236"
			return "1150"
		when "Z77", "Z75", "H77", "Q77", "B75", "Z68", "P67", "H67", "Q67", "B65", "H61"
			return "1155"
		when "X399"
			return "TR4"
		when "X470", "X370", "B450", "B350", "A320"
			return "AM4"
		when "990FX", "990X", "970", "890GX", "890FX", "880G", "870", "785G", "770", "760G", "GeForce 7025", "nForce 520LE"
			return "AM3+"
		when "A88X", "A85X", "A78", "A75", "A68H", "A58", "A55"
			return "FM2+"
		when "A50M", "A45"
			return "APU"
		when "NM70", "NM10"
			return "Intel CPU"
		else
			return "?"
		end
	end

	def write_file(pat)
		if pat == 0
			File.new("tmp/etailers.html", "w:UTF-8")
		elsif pat == 1
			File.open("tmp/etailers.html", 'a:UTF-8') {|file| file.write @browser.html}
		elsif pat == 2
			File.open("tmp/etailers.html", 'w:UTF-8') {|file| file.write @browser.html}
		end
	end

	def application_check_save(site, record) # record -> 是否記錄 (True / False)
		check_db = Site.find_by(site: site)
		if check_db
			if (Time.now - check_db.check_time.in_time_zone('Taipei')) > 28800 #8hours
				check_db.update(site: site, check_time: Time.now) if record
				return true
			end
		else
			Site.create(site: site, check_time: Time.now) if record
			return true
		end
		return false
	end

	def gen_cell(val, posx, posy, font_color, bgcolor, align, no_format, hyperlink)
		if hyperlink
			cell = @worksheet.add_cell(posx, posy, '', hyperlink)
		else
			cell = @worksheet.add_cell(posx, posy, val)
		end
		cell.set_number_format(no_format) if no_format
		cell.change_font_color color if font_color
		cell.change_fill bgcolor if bgcolor
		cell.change_horizontal_alignment(align) if align
		[:top, :bottom, :left, :right].each do |e|
			cell.change_border(e, "thin")
		end
	end

end
