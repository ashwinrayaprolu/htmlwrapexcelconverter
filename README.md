# htmlwrapexcelconverter
Everyone started facing issues with HTML Wrapped excel export with Latest KB from Microsoft. This will provide converter to fix that issue


This is a maven project setup as a WAR packaging, with an EmbedMe class in
the test scope that starts an embedded jetty of the WAR file being
produced by this project.

Quick Start
-----------

    $ mvn clean install exec:exec

Open your web browser to

    http://localhost:8080/convert?filePath=  to test Conversion
