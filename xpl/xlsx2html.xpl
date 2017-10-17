<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step 
  xmlns:p="http://www.w3.org/ns/xproc"
  xmlns:bc="http://transpect.le-tex.de/book-conversion"
  xmlns:c="http://www.w3.org/ns/xproc-step"
  xmlns:cx="http://xmlcalabash.com/ns/extensions" 
  xmlns:cxf="http://xmlcalabash.com/ns/extensions/fileutils"
  xmlns:tr="http://transpect.io"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  name="xlsx2html"
  type="tr:xlsx2html"
  version="1.0">
  
  <p:input port="params" kind="parameter" primary="true"/>

  <p:input port="xlsx2html-xsl">
    <p:document href="../xsl/xlsx2html.xsl"/>
  </p:input>
  <p:input port="html2csv-xsl">
    <p:document href="../xsl/html2csv.xsl"/>
  </p:input>

  <p:output port="result" primary="true">
    <p:pipe port="result" step="strip-srcpaths"/>
  </p:output>
  <p:serialization port="result" omit-xml-declaration="false" method="xhtml" indent="true"/>
  
  <p:output port="csv">
    <p:pipe port="result" step="transform-html2csv"/>
  </p:output>
  <p:serialization port="csv" method="text"/>

  <p:option name="in-file" required="true"/>
  <p:option name="debug" select="'no'"/> 
  <p:option name="debug-dir-uri" select="'debug'"/>
  <p:option name="out-dir-uri" select="''"/>
  <p:option name="cell-separator" select="'tab'"/>
  <p:option name="line-separator" select="'CRLF'"/>
  <p:option name="sheet-separator" select="'hyphens'"/>

  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl" />
  <p:import href="http://transpect.io/xproc-util/store-debug/xpl/store-debug.xpl"/>
  <p:import href="http://transpect.io/calabash-extensions/unzip-extension/unzip-declaration.xpl"/>
  <p:import href="http://transpect.io/xproc-util/xml-model/xpl/prepend-hub-xml-model.xpl" />
  <p:import href="http://transpect.io/xproc-util/file-uri/xpl/file-uri.xpl" />

  <p:variable name="basename" select="replace($in-file, '^(.+?)([^/\\]+)\.xlsx$', '$2')"/>

  <tr:file-uri name="file-uri">
    <p:with-option name="filename" select="$in-file"/>
  </tr:file-uri>

  <tr:store-debug>
    <p:with-option name="pipeline-step" select="concat(replace($in-file, '^.+/(.+)\.xlsx$', '$1.debug', 'i'), 
                                                       '/xlsx2html/00_file-uri')"/>
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </tr:store-debug>

  <tr:unzip name="xlsx-unzip">
    <p:with-option name="zip" select="/*/@os-path"/>
    <p:with-option name="dest-dir" select="concat(/*/@os-path, '.tmp')"/>
    <p:with-option name="overwrite" select="'yes'" />
    <p:documentation>Unzips the .xlsx file.</p:documentation>
  </tr:unzip>

  <p:xslt name="unzip">
    <p:input port="source">
      <p:pipe port="result" step="xlsx-unzip"/>
    </p:input>
    <p:input port="stylesheet">
      <p:inline>
        <xsl:stylesheet version="2.0">
          <xsl:template match="*|@*">
            <xsl:copy>
              <xsl:apply-templates select="@*, node()"/>
            </xsl:copy>
          </xsl:template>
          <xsl:template match="@name">
            <xsl:attribute name="name" select="replace(replace(., '\[', '%5B'), '\]', '%5D')"/>
          </xsl:template>
        </xsl:stylesheet>
      </p:inline>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
  </p:xslt>

  <tr:store-debug>
    <p:with-option name="pipeline-step" select="concat(replace($in-file, '^.+/(.+)\.xlsx$', '$1.debug', 'i'), 
                                                       '/xlsx2html/00_unzip-filelist')"/>
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </tr:store-debug>

  <p:group name="merge">
    <p:output port="result" primary="true"/>
    <p:variable name="base-dir" select="/c:files/@xml:base">
      <p:pipe port="result" step="unzip"/>
    </p:variable>
    <p:for-each>
      <p:iteration-source select="/c:files/c:file[not(matches(@name, '\.(bin|jpe?g|vml|png)$'))]"/>
      <p:choose>
        <p:when test="$debug = 'yes'">
          <cx:message>
            <p:with-option name="message" select="concat('xlsx2html info: Loading file &quot;', $base-dir, /c:file/@name, '&quot;')"/>
          </cx:message>    
        </p:when>
        <p:otherwise>
          <p:identity/>
        </p:otherwise>
      </p:choose>
      <p:load>
        <p:with-option name="href" select="concat($base-dir, /c:file/@name)"/>
      </p:load>
      <p:wrap wrapper="part" match="/*"/>
      <p:add-xml-base/>
    </p:for-each>
    <p:wrap-sequence wrapper="xlsx-parts"/>
    <p:add-attribute attribute-name="xml:base" match="/*">
      <p:with-option name="attribute-value" select="$base-dir"/>
    </p:add-attribute>
  </p:group>

  <p:xslt name="add-src-paths">
    <p:documentation>Adds src-path information to each element.</p:documentation>
    <p:input port="source">
      <p:pipe port="result" step="merge"/>
    </p:input>
    <p:input port="stylesheet">
      <p:document href="../xsl/add-src-paths.xsl"/>
    </p:input>
    <p:input port="parameters">
      <p:empty/>
    </p:input>
  </p:xslt>

  <tr:store-debug>
    <p:with-option name="pipeline-step" select="concat(replace($in-file, '^.+/(.+)\.xlsx$', '$1.debug', 'i'), 
                                                       '/xlsx2html/01_mergedParts')"/>
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </tr:store-debug>

  <p:sink/>

  <p:filter select="//*:propmap" name="filter-propmap">
    <p:input port="source">
      <p:document href="http://transpect.io/docx2hub/xsl/modules/prop-mapping/propmap.xsl"/>
    </p:input>
  </p:filter>

  <p:string-replace match="@*[matches(., '(^|\W)w:')]" replace="replace(replace(., '(^|\W)w:', '$1'), '^rFonts$', 'rFont')" name="transform-propmap"/>

  <tr:store-debug>
    <p:with-option name="pipeline-step" select="concat(replace($in-file, '^.+/(.+)\.xlsx$', '$1.debug', 'i'), 
                                                               '/xlsx2html/propmap.modified')"/>
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
    <p:with-option name="extension" select="'xsl'"/>
  </tr:store-debug>

  <p:sink/>

  <p:xslt template-name="main" name="transform-xlsx2html-main">
    <p:input port="source">
      <p:pipe port="result" step="add-src-paths"/>
      <p:pipe port="result" step="transform-propmap"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe port="xlsx2html-xsl" step="xlsx2html"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
  </p:xslt>

  <tr:store-debug>
    <p:with-option name="pipeline-step" select="concat(replace($in-file, '^.+/(.+)\.xlsx$', '$1.debug', 'i'), 
                                                       '/xlsx2html/02_xlsx2html')"/>
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </tr:store-debug>

  <p:delete match="@srcpath" name="strip-srcpaths"/>

  <p:xslt name="transform-html2csv">
    <p:input port="stylesheet">
      <p:pipe port="html2csv-xsl" step="xlsx2html"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
    <p:with-param name="cell-separator" select="$cell-separator"/>
    <p:with-param name="line-separator" select="$line-separator"/>
    <p:with-param name="sheet-separator" select="$sheet-separator"/>
  </p:xslt>

</p:declare-step>