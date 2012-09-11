import AssemblyKeys._

name := "scaxls"

version := "1.0.9"

organization := "jp.co.infocraft"

scalaVersion := "2.9.1"

resolvers ++= Seq(
  "twitter json Maven2 Repository" at "http://maven.twttr.com"
)

assemblySettings

mainClass in assembly := Some("jp.co.infocraft.scaxls.Main")

libraryDependencies ++= Seq(
		    "log4j" % "log4j" % "1.2.17",
		    "org.scalatest" %% "scalatest" % "1.8",
		    "junit" % "junit" % "4.10",
		    "org.apache.poi" % "poi" % "3.8",
		    "com.twitter" % "json_2.9.1" % "2.1.7"
		    )
