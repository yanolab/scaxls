package jp.co.infocraft.scaxls

import scala.io.Source
import java.io.InputStream
import java.io.FileInputStream
import java.io.File
import java.io.FileOutputStream
import java.io.OutputStream
import java.lang.System.exit

/**
 * Created by IntelliJ IDEA.
 * User: Yano
 * Date: 11/11/17
 * Time: 20:33
 * To change this template use File | Settings | File Templates.
 */

object Main {
  def main(args: Array[String]) = {
    var _type: String = "json"
    var _in: InputStream = System.in
    var _out: OutputStream = System.out
    var _template:String = null
    var _usage = false
    var _mode = "write"

    for (index <- 0 to (args.length-1)) {
      args(index) match {
        case "-type" => _type = args(index+1)
        case "-in" => _in = new FileInputStream(new File(args(index+1)))
        case "-out" => _out = new FileOutputStream(new File(args(index+1)))
        case "-template" => _template = args(index+1)
	case "-help" => _usage = true
	case "-mode" => _mode = args(index+1)
        case _ =>
      }
    }

    if(_usage) {
      println("java -jar scaxls.jar [-mode write|read(default:write)] [-type json] [-in filename(default:stdin)] [-out filename(default:stdout)] [-template filename] [-help]")
      exit(0);
    }

    val excel = new Excel(_template)

    _mode.toLowerCase match {
      case "write" => excel.render(Excel.parse(Source.fromInputStream(_in).getLines().mkString("\n"), _type), _out)
      case "read" => excel.read(_out)
    }

    _in.close()
    _out.close()
  }
}
