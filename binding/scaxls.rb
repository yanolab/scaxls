class Scaxls
  def self.type_and_value(val)
    if val.instance_of?(String) or val.is_a?(Numeric)
      return nil, val
    elsif val.instance_of?(Time) or val.instance_of?(DateTime)
      return "datetime", val.to_f * 1000.0
    elsif val.instance_of?(Date)
      return "date", val.to_time.to_f * 1000.0
    elsif val.instance_of?(TrueClass) or val.instance_of?(FalseClass)
      return nil, val.to_s
    else
      return nil, val
    end
  end

  def self.autoload
    here = File.expand_path(File.dirname(__FILE__))
    lib = Dir.glob(File.join(here, "scaxls-*.jar")).sort.last
    raise RuntimeException.new("library not found.") if lib.nil?
    lib
  en

  def self.json2xls(json_data, lib=autoload)
    IO.popen("java -jar #{lib}", "w+b") do |io|
      io.write(json_data)
      io.close_write
      io.read
    end
  end

  def self.xls2json(filename, lib=autoload)
    IO.popen("java -jar #{lib} -template #{filename} -mode read", "r") {|io| io.read}
  end
end
