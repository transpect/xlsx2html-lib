<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0"
                xmlns:xsl	= "http://www.w3.org/1999/XSL/Transform"
                xmlns:xs = "http://www.w3.org/2001/XMLSchema"
                xmlns:saxon	= "http://saxon.sf.net/"
                xmlns:letex	= "http://www.le-tex.de/namespace"
                xmlns:xlsx2html = "http://www.le-tex.de/namespace/xlsx2html"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns="http://www.w3.org/1999/xhtml"
                exclude-result-prefixes = "xs saxon letex"
                >
  
  <xsl:output
      method="xml"
      encoding="utf-8"
      indent="no"
      cdata-section-elements=''
      />

  <xsl:param name="base-uri" select="base-uri(/)"/>

  <xsl:variable name="base-dir" select="/*:xlsx-parts/@xml:base" as="xs:string"/>
  <xsl:variable name="main-rels-part" select="concat($base-dir, '_rels/.rels')" as="xs:string"/>
  <xsl:variable name="main-part" select="//*:part[@xml:base eq concat($base-dir, //*:part[@xml:base eq $main-rels-part]//*:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument']/@Target)]" as="node()"/>
  <xsl:variable name="main-part-dir" select="replace($main-part/@xml:base, '^(.+/)[^/]+$', '$1')" as="xs:string"/>
  <xsl:variable name="main-part-name" select="replace($main-part/@xml:base, '^(.+/)([^/]+)$', '$2')" as="xs:string"/>

  <xsl:template name="main">
    <xsl:apply-templates select="$main-part" mode="html"/>
  </xsl:template>

  <xsl:template match="*:part" mode="html">
    <xsl:for-each select="*:workbook">
      <xsl:element name="html">
        <xsl:element name="head">
          <xsl:element name="title">
            <xsl:value-of select="$base-dir"/>
          </xsl:element>
        </xsl:element>
        <xsl:element name="body">
          <xsl:apply-templates select="*:sheets" mode="#current"/>
        </xsl:element>
      </xsl:element>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="*:sheets" mode="html">
    <xsl:for-each select="*:sheet">
      <xsl:variable name="rels-file" select="//*:part[@xml:base eq concat($main-part-dir, '_rels/', $main-part-name, '.rels')]"/>
      <xsl:variable name="rel-target" select="concat($main-part-dir, $rels-file//*:Relationship[@Id eq current()/@r:id]/@Target)"/>
      <xsl:apply-templates select="//*:part[@xml:base eq $rel-target]/*:worksheet" mode="html">
        <xsl:with-param name="title" select="@name" tunnel="yes"/>
      </xsl:apply-templates>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="*:worksheet" mode="html">
    <xsl:param name="title" as="xs:string" tunnel="yes"/>
    <xsl:element name="div">
      <xsl:attribute name="class" select="local-name()"/>
      <xsl:element name="h2">
        <xsl:value-of select="$title"/>
      </xsl:element>
      <xsl:apply-templates select="*:sheetData" mode="#current"/>
    </xsl:element>
  </xsl:template>

  <xsl:template match="*:sheetData" mode="html">
    <xsl:variable name="context" select="." as="node()+"/>
    <xsl:variable name="mergedCellsInfo" as="node()*">
      <mergedCellsInfo>
        <xsl:for-each select="following-sibling::*/self::*:mergeCells/*:mergeCell">
          <mergedCell>
            <xsl:variable name="mergedCellStart" select="tokenize(@ref, ':')[1]"/>
            <xsl:variable name="mergedCellEnd" select="tokenize(@ref, ':')[2]"/>
            <xsl:variable name="mergedCellStartCol" select="replace($mergedCellStart, '^([A-Z]+)([0-9]+$)', '$1')"/>
            <xsl:variable name="mergedCellStartRow" select="replace($mergedCellStart, '^([A-Z]+)([0-9]+$)', '$2')"/>
            <xsl:variable name="mergedCellEndCol" select="replace($mergedCellEnd, '^([A-Z]+)([0-9]+$)', '$1')"/>
            <xsl:variable name="mergedCellEndRow" select="replace($mergedCellEnd, '^([A-Z]+)([0-9]+$)', '$2')"/>
            <xsl:variable name="mergedRows" select="for $i in (xs:integer($mergedCellStartRow) to xs:integer($mergedCellEndRow)) return $i"/>
            <xsl:variable name="mergedCols" select="for $r in $mergedRows return $context//*:c[   (@r eq concat($mergedCellStartCol, $r)) 
                                                                                               or (@r eq concat($mergedCellEndCol, $r))
                                                                                               or (preceding-sibling::*[@r eq concat($mergedCellStartCol, $r)] and following-sibling::*[@r eq concat($mergedCellEndCol, $r)])]/@r"/>
            <mergedCellRef><xsl:value-of select="@ref"/></mergedCellRef>
            <mergedCellStart><xsl:value-of select="$mergedCellStart"/></mergedCellStart>
            <mergedCellEnd><xsl:value-of select="$mergedCellEnd"/></mergedCellEnd>
            <mergedCellStartCol><xsl:value-of select="$mergedCellStartCol"/></mergedCellStartCol>
            <mergedCellStartRow><xsl:value-of select="$mergedCellStartRow"/></mergedCellStartRow>
            <mergedCellEndCol><xsl:value-of select="$mergedCellEndCol"/></mergedCellEndCol>
            <mergedCellEndRow><xsl:value-of select="$mergedCellEndRow"/></mergedCellEndRow>
            <mergedRows><xsl:value-of select="$mergedRows"/></mergedRows>
            <mergedCols><xsl:value-of select="$mergedCols"/></mergedCols>
          </mergedCell>
        </xsl:for-each>
      </mergedCellsInfo>
    </xsl:variable>
    <xsl:element name="table">
      <xsl:apply-templates select="@*|node()" mode="#current">
        <xsl:with-param name="mergedCellsInfo" select="$mergedCellsInfo" tunnel="yes"/>
      </xsl:apply-templates>
    </xsl:element>
  </xsl:template>

  <xsl:template match="*:row" mode="html">
    <xsl:param name="title" tunnel="yes"/>
    <xsl:element name="tr">
      <xsl:attribute name="id" select="concat($title, '_ROW', @r, '_COLS', replace(@spans, ':', '-'))"/>
      <xsl:apply-templates select="@*|node()" mode="#current"/>
    </xsl:element>
  </xsl:template>

  <xsl:template match="*:row/@r" mode="html"/>
  <xsl:template match="*:row/@spans" mode="html"/>

  <xsl:template match="*:c" mode="html">
    <xsl:param name="mergedCellsInfo" tunnel="yes" as="node()*"/>
    <xsl:choose>
      <xsl:when test="every $a in xlsx2html:mergedCells($mergedCellsInfo, .) satisfies ($a/local-name() eq 'MERGEDCELL')">
        <xsl:if test="normalize-space(.)">
          <xsl:message select="concat('xlsx2html Warning: Content found in deleted merged cell. This should not happen. Context: ', .)"/>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <xsl:element name="td">
          <xsl:sequence select="xlsx2html:mergedCells($mergedCellsInfo, .)"/>
          <xsl:apply-templates select="@*|node()" mode="#current"/>
        </xsl:element>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:function name="xlsx2html:mergedCells">
    <xsl:param name="mergedCellsInfo" as="node()*"/>
    <xsl:param name="currentCell" as="node()"/>
    <xsl:choose>
      <xsl:when test="some $i in $mergedCellsInfo//*:mergedCellStart satisfies ($i eq $currentCell/@r)"><!-- current cell is starting merged cell -->
        <xsl:variable name="info-tag" select="$mergedCellsInfo/*:mergedCell[*:mergedCellStart eq $currentCell/@r]" as="node()*"/>
        <xsl:attribute name="rowspan" select="count(tokenize($info-tag/*:mergedRows, ' '))"/>
        <xsl:attribute name="colspan" select="count($currentCell/following-sibling::*:c[following-sibling::*:c[@r eq concat($info-tag/*:mergedCellEndCol, $info-tag/*:mergedCellStartRow)]])+2"/>
      </xsl:when>
      <xsl:when test="some $i in tokenize(string-join($mergedCellsInfo//*:mergedCols, ' '), ' ') satisfies ($i eq $currentCell/@r)"><!-- current cell is part of a merged cell and should be deleted -->
        <xsl:attribute name="MERGEDCELL" select="'yes'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="rowspan" select="'1'"/>
        <xsl:attribute name="colspan" select="'1'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template match="*:c/@s" mode="html"/><!-- may be analyzed later -->

  <xsl:template match="*:c/@r" mode="html">
    <xsl:attribute name="axis" select="."/>
  </xsl:template>

  <xsl:template match="*:c/@t" mode="html"><!--type: boolean, number, string-->
    <xsl:attribute name="class" select="."/>
  </xsl:template>

  <xsl:template match="*:v" mode="html">
    <xsl:choose>
      <xsl:when test="parent::*/@t='s'"><!-- string value: sharedStringTable lookup! -->
        <xsl:variable name="string-from-sst" select="xlsx2html:get-string-from-sst(., root()//*:sst)" as="node()*"/>
        <xsl:apply-templates select="$string-from-sst" mode="#current"/>
      </xsl:when>
      <xsl:when test="preceding-sibling::*:f or following-sibling::*:f"><!-- formula: value is the result of the most recent calculation -->
        <xsl:if test="count(preceding-sibling::*:f | following-sibling::*:f) gt 1">
          <xsl:message select="concat('xlsx2html Warning: More than one formula found: ', parent::*)"/>
        </xsl:if>
        <xsl:element name="span">
          <xsl:attribute name="class" select="'formula'"/>
          <xsl:apply-templates select="preceding-sibling::*:f | following-sibling::*:f" mode="#current">
            <xsl:with-param name="render" select="'yes'"/>
          </xsl:apply-templates>
        </xsl:element>
        <xsl:element name="span">
          <xsl:attribute name="class" select="'result'"/>
          <xsl:apply-templates select="node()" mode="#current"/>
        </xsl:element>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="node()" mode="#current"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:function name="xlsx2html:get-string-from-sst" as="node()*">
    <xsl:param name="index" as="xs:integer"/>
    <xsl:param name="sst" as="node()"/>
    <xsl:sequence select="$sst/*:si[position() eq ($index+1)]//*:t/node()"/><!--Include inline formatting (<r>, <rPr>) here! You may adapt this from wml2dbk, if possible.-->
  </xsl:function>

  <xsl:template match="*:f" mode="html">
    <xsl:param name="render" select="'no'"/>
    <xsl:if test="$render='yes'">
      <xsl:value-of select="."/>
    </xsl:if>
  </xsl:template>



  <!-- catch-all -->

  <xsl:template match="*" mode="#all">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:apply-templates select="processing-instruction() | node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="node()|@*" priority="-1" mode="#all">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:apply-templates select="node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>


</xsl:stylesheet>