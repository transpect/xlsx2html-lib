<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step 
  xmlns:p="http://www.w3.org/ns/xproc"
  xmlns:bc="http://transpect.le-tex.de/book-conversion"
  xmlns:c="http://www.w3.org/ns/xproc-step"
  xmlns:cx="http://xmlcalabash.com/ns/extensions" 
  xmlns:cxf="http://xmlcalabash.com/ns/extensions/fileutils"
  xmlns:transpect="http://www.le-tex.de/namespace/transpect"
  xmlns:letex="http://www.le-tex.de/namespace"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  name="xlsx2html"
  version="1.0">
  
  <p:input port="params" kind="parameter" primary="true">
    <p:documentation>Arbitrary parameters that will be passed to the dynamically executed pipeline.</p:documentation>
  </p:input>

  <p:input port="stylesheet">
    <p:document href="../xsl/xlsx2html.xsl"/>
  </p:input>

  <p:output port="result" primary="false" sequence="true"/>

  <p:option name="in-file" required="true"/>
  <p:option name="debug" select="'no'"/> 
  <p:option name="debug-dir-uri" select="'debug'"/>
  <p:option name="out-dir-uri" select="''"/>
  
  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl" />
  <p:import href="http://transpect.le-tex.de/xproc-util/store-debug/store-debug.xpl"/>
  <p:import href="http://transpect.le-tex.de/calabash-extensions/ltx-lib.xpl" />
  <p:import href="http://transpect.le-tex.de/xproc-util/xml-model/prepend-hub-xml-model.xpl" />

  <p:variable name="basename" select="replace($in-file, '^(.+?)([^/\\]+)\.xlsx$', '$2')"/>

  <letex:unzip name="xlsx-unzip">
    <p:with-option name="zip" select="$in-file" />
    <p:with-option name="dest-dir" select="concat($in-file, '.tmp')"><p:empty/></p:with-option>
    <p:with-option name="overwrite" select="'yes'" />
    <p:documentation>Unzips the .xlsx file.</p:documentation>
  </letex:unzip>

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

  <letex:store-debug pipeline-step="00_unzip-filelist">
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </letex:store-debug>

  <p:group name="merge">
    <p:output port="result" primary="true"/>
    <p:variable name="base-dir" select="/c:files/@xml:base">
      <p:pipe port="result" step="unzip"/>
    </p:variable>
    <p:for-each>
      <p:iteration-source select="/c:files/c:file[not(matches(@name, '\.bin$'))]"/>
      <cx:message>
        <p:with-option name="message" select="concat('xlsx2html info: Loading file &quot;', $base-dir, /c:file/@name, '&quot;')"/>
      </cx:message>
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

  <letex:store-debug pipeline-step="01_mergedParts">
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </letex:store-debug>

  <p:sink/>

  <p:xslt template-name="main">
    <p:input port="source">
      <p:pipe port="result" step="merge"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe port="stylesheet" step="xlsx2html"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
  </p:xslt>

  <letex:store-debug pipeline-step="02_xslt">
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </letex:store-debug>

  <p:store>
    <p:with-option name="href" select="concat($out-dir-uri, $basename, '.html')"/>
  </p:store>

</p:declare-step>