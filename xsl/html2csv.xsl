<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0"
                xmlns:css = "http://www.w3.org/1996/css"
                xmlns:xsl	= "http://www.w3.org/1999/XSL/Transform"
                xmlns:xs = "http://www.w3.org/2001/XMLSchema"
                xmlns:tr	= "http://transpect.io"
                xmlns:html = "http://www.w3.org/1999/xhtml"
                xpath-default-namespace="http://www.w3.org/1999/xhtml"
                xmlns="http://www.w3.org/1999/xhtml"
                exclude-result-prefixes = "xs tr css"
                >

  <xsl:param name="csv-separator" select="'&#x9;'" as="xs:string"/>
  <xsl:param name="line-separator" select="'&#xa;'" as="xs:string"/>

  <xsl:template match="text()" mode="#default"/>
  
  <xsl:template match="body" mode="#default">
    <csv>
      <xsl:apply-templates mode="csv"/>
    </csv>
  </xsl:template>
  
  <xsl:template match="h1 | h2 | h3 | h4 | h5 | h6 | p" mode="csv">
    <xsl:value-of select="normalize-space()"/>
    <xsl:value-of select="$line-separator"/>
  </xsl:template>
  
  <xsl:template match="span[@class = 'formula']" mode="csv"/>
  
  <xsl:template match="tr" mode="csv">
    <xsl:variable name="cells" as="xs:string*">
      <xsl:apply-templates select="*" mode="#current"/>
    </xsl:variable>
    <xsl:value-of select="string-join($cells, $csv-separator)"/>
    <xsl:value-of select="$line-separator"/>
  </xsl:template>
  
  <xsl:template match="td | th" mode="csv" as="text()*">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>
  
  <xsl:template match="td[not(normalize-space())] | th[not(normalize-space())]" mode="csv" as="text()*">
    <xsl:text/>
  </xsl:template>
  

</xsl:stylesheet>