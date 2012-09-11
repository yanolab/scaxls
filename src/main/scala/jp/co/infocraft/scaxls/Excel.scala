package jp.co.infocraft.scaxls

import java.io.{File, FileInputStream, FileOutputStream, OutputStream}
import java.util.Date

import scala.collection.mutable.{HashMap}

import org.apache.poi.hssf.usermodel.{HSSFDataFormat, HSSFWorkbook}
import org.apache.poi.hssf.util.HSSFCellUtil
import org.apache.poi.ss.usermodel.{Cell, CellStyle}

import com.twitter.json.Json

/**
 * Created by IntelliJ IDEA.
 * User: Yano
 * Date: 11/11/18
 * Time: 11:33
 * To change this template use File | Settings | File Templates.
 */

class Excel(template:String = null) {
  private val _workbook = template match {
    case null => new HSSFWorkbook
    case _ => new HSSFWorkbook(new FileInputStream(template))
  }

  private val _defaultCellStyle = _workbook.createCellStyle
  _defaultCellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("text"))

  def read(outstream: OutputStream) {
    var rdata = List[Map[String, Any]]()

    for(idx <- 0 to (_workbook.getNumberOfSheets-1)) {
      val sheet = _workbook.getSheetAt(idx)
      var cells = List[Map[String,Any]]()

      for(rowidx <- sheet.getFirstRowNum to sheet.getLastRowNum;
	val row = sheet.getRow(rowidx)
	if row != null)
      {
	for(cellidx <- row.getFirstCellNum to row.getLastCellNum;
	    val cell = row.getCell(cellidx)
	    if cell != null)
	{
	  val cellval = cell.getCellType match {
	    case Cell.CELL_TYPE_NUMERIC => cell.getNumericCellValue
	    case _ => cell.getStringCellValue
	  }

	  val cellmap = Map("x" -> cell.getColumnIndex,
			    "y" -> cell.getRowIndex,
			    "value" -> cellval)

	  cells = cellmap :: cells
	}
      }
      rdata = Map("name" -> sheet.getSheetName, "cells" -> cells) :: rdata
    }

    outstream.write(Json.build(rdata).toString.getBytes)
  }

  def render(sheets: List[Map[String, Any]], outstream: OutputStream) {
    val styleMap = new HashMap[String, CellStyle]

    def _getInt(v: Any) = v match {
      case Some(v: Int) => v
      case _ => -1
    }

    for (sheet <- sheets) {
      val sheetname:Any = sheet.get("name")
      val worksheet = sheetname match {
        case Some(v:String) => _workbook.getSheet(v) match {
          case null => _workbook.createSheet(v)
          case v => v
        }
        case _ => _workbook.createSheet
      }

      for (
        cellmap <- sheet("cells").asInstanceOf[List[Map[String, Any]]];
        val x = _getInt(cellmap.get("x"));
        val y = _getInt(cellmap.get("y"));
        if x >= 0 && y >= 0
      ) {
        val row = HSSFCellUtil.getRow(y, worksheet)
        val cell = HSSFCellUtil.getCell(row, x)

        (cellmap.get("type"), cellmap.get("format"), cellmap.get("value")) match {
          case (Some("formula"), _, Some(v: String)) => {
            cell.setCellType(Cell.CELL_TYPE_FORMULA)
            cell.setCellFormula(v)
          }
          case (_, _, Some(v: String)) => {
            cell.setCellType(Cell.CELL_TYPE_STRING)
            cell.setCellValue(v)
            cell.setCellStyle(_defaultCellStyle)
          }
          case (Some(tp: String), format, Some(v: Number)) if (tp == "datetime" || tp == "date") => {
            cell.setCellValue(new Date(v.longValue()))

	    val formatString = format match {
	        case Some(v: String) => v
                case _ => tp match {
	      	    case "datetime" => "yyyy/mm/dd HH:MM"
		    case "date" => "yyyy/mm/dd"
	        }

	    }

	    val style = styleMap.get(formatString) match {
	        case Some(v) => v
		case _ => {
		     val newStyle = _workbook.createCellStyle()
            	     newStyle.setDataFormat(_workbook.createDataFormat().getFormat(formatString))
		     styleMap.put(formatString, newStyle)
		     newStyle
            	}
            }

            cell.setCellStyle(style)
          }
          case (_, _, Some(v: Number)) => {
            cell.setCellType(Cell.CELL_TYPE_NUMERIC)
            cell.setCellValue(v.doubleValue())
          }
          case (_, _, Some(null)) => {
	       // NOP
          }
          case v => {
            Console.err.println("X:%d, Y:%d, Unmatch Data: %s" format (x, y, v.toString()))
          }
        }
      }
    }

    save(outstream)
  }

  /**
   *
   */
  def save(file: File) {
    val fos = new FileOutputStream(file)
    save(fos)
    fos.close()
  }

  def save(outstream:OutputStream) {
    _workbook.write(outstream)
  }

  /**
   * *
   *
   */
  def save(filename: String) {
    this.save(new File(filename))
  }
}

object Excel {
  def parse(in: String, dataType: String): List[Map[String, Any]] = {
    dataType match {
      case "xml" => parseXML(in)
      case "json" => parseJSON(in)
      case "mpack" => parseMPACK(in)
      case v => throw new RuntimeException("Unkwon data type %s" format v)
    }
  }

  def parseJSON(in: String): List[Map[String, Any]] = {
    return Json.parse(in).asInstanceOf[List[Map[String, Any]]]
  }

  def parseMPACK(in: String): List[Map[String, Any]] = {
    throw new RuntimeException("MPack Parser Not implement.")
  }

  def parseXML(in: String): List[Map[String, Any]] = {
    throw new RuntimeException("XML Parser Not implement.")
  }

}
