# htmlwrapexcelconverter
Everyone started facing issues with HTML Wrapped excel export with Latest KB from Microsoft. This will provide converter to fix that issue

Took base Embedded Jetty code from 

    https://github.com/jetty-project/embedded-servlet-3.1

This is a maven project setup as a WAR packaging, with an EmbedMe class in
the test scope that starts an embedded jetty of the WAR file being
produced by this project.

Quick Start
-----------

    $ mvn clean install exec:exec

Open your web browser to

    http://localhost:10100/convert?filePath=/Users/ashwinrayaprolu/Desktop/AuditCheck.xls  to test Conversion
