require "google_drive"

session = GoogleDrive::Session.from_config("config.json")
ws = session.spreadsheet_by_key("1gEcniimJD-V61-waYqO8kuGWP1XoqvFHzLDuJA9-Ki8").worksheets[0]

class GoogleSheet

  attr_reader :worksheet

  def initialize(worksheet)
    @worksheet = worksheet
  end

  def method_missing(method_name, *arguments, &block)
    if method_name.to_s.match?(/\A[a-zA-Z]+Kolona\z/)
      column_index = method_name.to_s.gsub(/[^\d]/, '').to_i
      Column.new(@worksheet, column_index)
    else
      super
    end
  end

  def find_row_by_cell_value(column_index, value)
    @worksheet.rows.find { |row| row[column_index - 1] == value }
  end

  def respond_to_missing?(method_name, include_private = false)
    method_name.to_s.end_with?('Kolona') || super
  end

  def empty_rows
    @worksheet.rows.select { |row| row.all?(&:empty?) }
  end

  def +(other_sheet)
    raise ArgumentError, 'Headeri oba sheet-a moraju biti isti!' unless headers_match?(other_sheet)

    combined_rows = @worksheet.rows + other_sheet.worksheet.rows[1..]
    GoogleSheet.new_from_rows(combined_rows)
  end

  def self.new_from_rows(rows)
    # Simulacija worksheet-a koristeći Struct.
    # Ovo služi kao privremeni "worksheet" za novi GoogleSheet objekat.
    simulated_worksheet = Struct.new(:rows).new(rows)

    # Kreiranje novog GoogleSheet objekta sa simuliranim worksheet-om.
    GoogleSheet.new(simulated_worksheet)
  end

  private

  def headers_match?(other_sheet)
    @worksheet.rows.first == other_sheet.worksheet.rows.first
  end

  def -(other_sheet)
    raise ArgumentError, 'Headeri oba sheet-a moraju biti isti!' unless headers_match?(other_sheet)

    new_rows = @worksheet.rows - other_sheet.worksheet.rows[1..]
    GoogleSheet.new_from_rows(new_rows)
  end

  class Column
    def initialize(worksheet, column_index)
      @worksheet = worksheet
      @column_index = column_index
    end

    def cells
      @worksheet.rows.reject { |row| row.any? { |cell| cell =~ /total|subtotal/i } }.map do |row|
        cell_value = row[@column_index - 1]
        if cell_value && !cell_value.empty? && cell_value.match?(/^\d+(\.\d+)?$/)
          cell_value.to_f
        else
          nil
        end
      end.compact
    end
    def sum
      cells.sum
    end

    def avg
      return 0 if cells.empty?
      cells.sum / cells.size
    end

    def map(&block)
      cells.map(&block)
    end

    def select(&block)
      cells.select(&block)
    end

    def reduce(initial, &block)
      cells.reduce(initial, &block)
    end
  end
end


def all_rows(ws)
  ws.rows.reject { |row| row.any? { |cell| cell =~ /total|subtotal/i } }
end

def row(ws, index)
  ws.rows[index]
end

def each_cell(ws)
  return enum_for(:each_cell, ws) unless block_given?
  ws.rows.each do |row|
    row.each do |cell|
      yield cell
    end
  end
end

def main(ws, session)

  ws1 = session.spreadsheet_by_key("1gEcniimJD-V61-waYqO8kuGWP1XoqvFHzLDuJA9-Ki8").worksheets[0]
  ws2 = session.spreadsheet_by_key("1gEcniimJD-V61-waYqO8kuGWP1XoqvFHzLDuJA9-Ki8").worksheets[1]

  sheet = GoogleSheet.new(ws)

  # Ispisuje prazne redove
  puts "Prazni redovi:"
  p sheet.empty_rows

  sheet = GoogleSheet.new(ws)
  puts "All Rows:"
  p all_rows(ws)

  puts "\nSpecific Row (Row 2):"
  p row(ws, 1)

  puts "\nIterating through each cell:"
  each_cell(ws) { |cell| print "#{cell} " }
  puts

  puts "\nSum of first column:"
  puts sheet.prvaKolona.sum

  puts "\nAverage of first column:"
  puts sheet.prvaKolona.avg

  puts "\nFind row by cell value:"
  p sheet.find_row_by_cell_value(1, "rn2310") # Pretpostavljamo da je 'rn2310' vrednost u prvoj koloni

  puts "\nMapping over first column:"
  sheet.prvaKolona.map { |cell| cell += 1 }

  sheet1 = GoogleSheet.new(ws1)
  sheet2 = GoogleSheet.new(ws2)

  combined_sheet = sheet1 + sheet2
  puts "Spojene tabele:"
  p combined_sheet.all_rows

  subtracted_sheet = sheet1 - sheet2
  puts "Oduzete tabele:"
  p subtracted_sheet.all_rows

end

main(ws, session)
