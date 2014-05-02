excel-splitter
==============

* sudo apt-get install maven2 openjdk7-jdk openjdk7-jre
* git clone https://github.com/Kagee/excel-splitter.git
* cd excel-splitter
* mvn package
* java -jar <path til excelsplit-1.0-SNAPSHOT-jar-with-dependencies.jar>

Missing code around line 142  and 149 in ES.java

* Konfigured for .xslx
 * For .xsl (.xls? ), change xsls to xsl in filenames, 
and change all instances of XSSFWorkbook to HSSFWorkbook

* Recommended editor
 * http://eclipse.org/downloads/packages/eclipse-ide-java-developers/keplersr2
 * File->import -> Maven -> Exsisting Maven Projects -> NEXT -> choose source folder -> FINISH

