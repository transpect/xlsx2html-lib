<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0"
                xmlns:css = "http://www.w3.org/1996/css"
                xmlns:xsl	= "http://www.w3.org/1999/XSL/Transform"
                xmlns:xs = "http://www.w3.org/2001/XMLSchema"
                xmlns:tr	= "http://transpect.io"
                xmlns:xlsx2html = "http://transpect.io/xlsx2html"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:xls="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
                xmlns="http://www.w3.org/1999/xhtml"
                exclude-result-prefixes = "xs tr xlsx2html x14ac xls css r rels"
                >

  <xsl:output
      method="xml"
      encoding="utf-8"
      indent="no"
      cdata-section-elements=''
      />

  <xsl:param name="base-uri" select="base-uri(/)"/>

  <xsl:variable name="propmap" select="collection()[2]" as="node()"/>

  <xsl:variable name="base-dir" select="/*:xlsx-parts/@xml:base" as="xs:string"/>
  <xsl:variable name="main-rels-part" select="concat($base-dir, '_rels/.rels')" as="xs:string"/>
  <xsl:variable name="main-part" as="element(part)?" 
    select="//part[@xml:base eq concat($base-dir, //part[@xml:base eq $main-rels-part]
              //rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument']/@Target)]"/>
  <xsl:variable name="main-part-dir" select="replace($main-part/@xml:base, '^(.+/)[^/]+$', '$1')" as="xs:string"/>
  <xsl:variable name="main-part-name" select="replace($main-part/@xml:base, '^(.+/)([^/]+)$', '$2')" as="xs:string"/>


  <xsl:template name="main">
    <xsl:apply-templates select="$main-part" mode="html"/>
  </xsl:template>

  <xsl:template match="*:part" mode="html">
    <xsl:for-each select="xls:workbook">
      <html>
        <xsl:apply-templates select="@srcpath" mode="#current"/>
        <head>
          <title>
            <xsl:value-of select="$base-dir"/>
          </title>
          <style type="text/css">
            span.formula { display: none }
          </style>
        </head>
        <body>
          <xsl:apply-templates select="xls:sheets" mode="#current"/>
        </body>
      </html>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="xls:sheets" mode="html">
    <xsl:for-each select="xls:sheet">
      <xsl:variable name="rels-file" select="//*:part[@xml:base eq concat($main-part-dir, '_rels/', $main-part-name, '.rels')]"/>
      <xsl:variable name="rel-target" select="concat($main-part-dir, $rels-file//*:Relationship[@Id eq current()/@r:id]/@Target)"/>
      <xsl:apply-templates select="//*:part[@xml:base eq $rel-target]/xls:worksheet" mode="html">
        <xsl:with-param name="title" select="@name" tunnel="yes"/>
        <xsl:with-param name="sheet-id" select="replace(replace(@name, '\C', '_'), '^(\I)', '_$1')" tunnel="yes"/>
      </xsl:apply-templates>
    </xsl:for-each>
  </xsl:template>


  <xsl:template match="xls:worksheet" mode="html">
    <xsl:param name="title" as="xs:string" tunnel="yes"/>
    <xsl:element name="div">
      <xsl:apply-templates select="@srcpath" mode="#current"/>
      <xsl:attribute name="class" select="local-name()"/>
      <xsl:element name="h2">
        <xsl:value-of select="$title"/>
      </xsl:element>
      <xsl:apply-templates select="xls:sheetData" mode="#current"/>
    </xsl:element>
  </xsl:template>

  <xsl:template match="xls:sheetData" mode="html">
    <xsl:variable name="context" select="." as="element(xls:sheetData)"/>
    <xsl:variable name="mergedCellsInfo" as="element(mergedCellsInfo)">
      <mergedCellsInfo xmlns="">
        <xsl:for-each select="following-sibling::*/self::xls:mergeCells/xls:mergeCell">
          <mergedCell>
            <xsl:variable name="mergedCellStart" select="tokenize(@ref, ':')[1]"/>
            <xsl:variable name="mergedCellEnd" select="tokenize(@ref, ':')[2]"/>
            <xsl:variable name="mergedCellStartCol" select="replace($mergedCellStart, '^([A-Z]+)([0-9]+$)', '$1')"/>
            <xsl:variable name="mergedCellStartRow" select="replace($mergedCellStart, '^([A-Z]+)([0-9]+$)', '$2')"/>
            <xsl:variable name="mergedCellEndCol" select="replace($mergedCellEnd, '^([A-Z]+)([0-9]+$)', '$1')"/>
            <xsl:variable name="mergedCellEndRow" select="replace($mergedCellEnd, '^([A-Z]+)([0-9]+$)', '$2')"/>
            <xsl:variable name="mergedRows" select="for $i in (xs:integer($mergedCellStartRow) to xs:integer($mergedCellEndRow)) return $i"/>
            <xsl:variable name="mergedCols" 
              select="for $r in $mergedRows return $context//xls:c[   (@r eq concat($mergedCellStartCol, $r)) 
                                                                   or (@r eq concat($mergedCellEndCol, $r))
                                                                   or (preceding-sibling::*[@r eq concat($mergedCellStartCol, $r)] 
                                                                       and following-sibling::*[@r eq concat($mergedCellEndCol, $r)])
                                                                  ]/@r"/>
            <xsl:variable name="colspan" select="xlsx2html:col-number($mergedCellEndCol) - xlsx2html:col-number($mergedCellStartCol) + 1"/>
            <xsl:variable name="rowspan" select="xs:integer($mergedCellEndRow) - xs:integer($mergedCellStartRow) + 1"/>
            <xsl:attribute name="colspan" select="$colspan[. gt 1]"/>
            <xsl:attribute name="rowspan" select="$rowspan[. gt 1]"/>
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

  <xsl:template match="xls:row" mode="html">
    <xsl:param name="sheet-id" tunnel="yes"/>
    <tr>
      <xsl:apply-templates select="@srcpath" mode="#current"/>
      <xsl:attribute name="id" select="concat($sheet-id, '_ROW', @r, '_COLS', replace(@spans, ':', '-'))"/>
      <xsl:apply-templates select="@*|node()" mode="#current"/>
    </tr>
  </xsl:template>

  <xsl:template match="xls:row/@*" mode="html"/>

  <xsl:template match="xls:c" mode="html">
    <xsl:param name="mergedCellsInfo" tunnel="yes" as="element(mergedCellsInfo)"/>
    <xsl:variable name="mergedcell" as="attribute(*)*" select="xlsx2html:mergedCells($mergedCellsInfo, .)"/>
    <xsl:choose>
      <xsl:when test="exists($mergedcell)
                      and
                      (
                        every $a in $mergedcell satisfies ($a/local-name() eq 'MERGEDCELL')
                      )">
        <xsl:if test="normalize-space(.)">
          <xsl:message select="concat('xlsx2html Warning: Content found in deleted merged cell. This should not happen. Context: ', .)"/>
        </xsl:if>
        <xsl:processing-instruction name="merged" select="@r"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="last-col" as="xs:integer" 
          select="xlsx2html:col-number(replace(preceding-sibling::xls:c[1]/@r, '\d', ''))"/>
        <xsl:variable name="current-col" as="xs:integer" 
          select="xlsx2html:col-number(replace(@r, '\d', ''))"/>
        <xsl:for-each select="($last-col + 1 to $current-col - 1)">
          <td class="filler"/>
        </xsl:for-each>
        <td>
          <xsl:apply-templates select="@srcpath" mode="#current"/>
          <xsl:sequence select="$mergedcell"/>
          <xsl:apply-templates select="@*|node()" mode="#current"/>
        </td>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:function name="xlsx2html:mergedCells" as="attribute(*)*">
    <xsl:param name="mergedCellsInfo" as="element(mergedCellsInfo)"/>
    <xsl:param name="currentCell" as="element(xls:c)"/>
    <xsl:variable name="span-atts" as="attribute(*)*" 
      select="$mergedCellsInfo/mergedCell[mergedCellStart eq $currentCell/@r]/@*[normalize-space()]"/><!-- colspan, rowspan -->
    <xsl:sequence select="$span-atts"/>
    <xsl:if test="empty($span-atts)
                  and 
                  (
                    some $i in tokenize(string-join($mergedCellsInfo//mergedCols, ' '), ' ') 
                    satisfies ($i eq $currentCell/@r)
                  )"><!-- current cell is part of a merged cell and should be deleted -->
      <xsl:attribute name="MERGEDCELL" select="'yes'"/>
    </xsl:if>
  </xsl:function>
  
  <xsl:function name="xlsx2html:col-number" as="xs:integer">
    <xsl:param name="col-chars" as="xs:string"/>
    <xsl:variable name="ints" as="xs:integer*">
      <xsl:for-each select="reverse(string-to-codepoints($col-chars))">
        <xsl:sequence select="(. - 64) * xlsx2html:pow(26, position() - 1)"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:sequence select="sum($ints)"/>
  </xsl:function>

  <xsl:function name="xlsx2html:pow" as="xs:integer">
    <xsl:param name="val" as="xs:integer"/>
    <xsl:param name="pow" as="xs:integer"/>
    <xsl:sequence select="if ($pow = 0) then 1 else $val * xlsx2html:pow($val, $pow - 1)"/>
  </xsl:function>

  <xsl:template match="xls:c/@s" mode="html"/><!-- may be analyzed later -->

  <xsl:template match="xls:c/@r" mode="html">
    <xsl:param name="sheet-id" as="xs:string" tunnel="yes"/>
    <xsl:attribute name="id" select="string-join(($sheet-id, .), '_')"/>
  </xsl:template>

  <xsl:template match="xls:c/@t" mode="html"><!--type: boolean, number, string-->
    <xsl:attribute name="class" select="."/>
  </xsl:template>
  
  <xsl:template match="@x14ac:*" mode="html"/>

  <xsl:template match="xls:v" mode="html">
    <xsl:choose>
      <xsl:when test="parent::*/@t='s'"><!-- string value: sharedStringTable lookup! -->
        <xsl:variable name="string-from-sst" select="xlsx2html:get-string-from-sst(., root()//xls:sst)" as="node()*"/>
        <xsl:apply-templates select="$string-from-sst" mode="#current"/>
      </xsl:when>
      <xsl:when test="preceding-sibling::xls:f or following-sibling::xls:f"><!-- formula: value is the result of the most recent calculation -->
        <xsl:if test="count(preceding-sibling::xls:f | following-sibling::xls:f) gt 1">
          <xsl:message select="concat('xlsx2html Warning: More than one formula found: ', parent::*)"/>
        </xsl:if>
        <xsl:element name="span">
          <xsl:apply-templates select="preceding-sibling::xls:f/@srcpath | following-sibling::xls:f/@srcpath" mode="#current"/>
          <xsl:attribute name="class" select="'formula'"/>
          <xsl:apply-templates select="preceding-sibling::xls:f | following-sibling::xls:f" mode="#current">
            <xsl:with-param name="render" select="'yes'"/>
          </xsl:apply-templates>
        </xsl:element>
        <xsl:element name="span">
          <xsl:apply-templates select="@srcpath" mode="#current"/>
          <xsl:attribute name="class" select="'result'"/>
          <xsl:value-of select="xlsx2html:apply-format(., ../@s)"/>
        </xsl:element>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="xlsx2html:apply-format(., ../@s)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Predefined number formats 
 https://social.msdn.microsoft.com/Forums/office/en-US/e27aaf16-b900-4654-8210-83c5774a179c/xlsx-numfmtid-predefined-id-14-doesnt-match?forum=oxmlsdk#4382183b-59a4-4863-9ad6-01470e9afcd9
 0 General
 1 0
 2 0.00
 3 #,##0
 4 #,##0.00
 5 $#,##0_);($#,##0)
 6 $#,##0_);[Red]($#,##0)
 7 $#,##0.00_);($#,##0.00)
 8 $#,##0.00_);[Red]($#,##0.00)
 9 0%
10 0.00%
11 0.00E+00
12 # ?/?
13 # ??/??
14 mm-dd-yy
15 d-mmm-yy
16 d-mmm
17 mmm-yy
18 h:mm AM/PM
19 h:mm:ss AM/PM
20 h:mm
21 h:mm:ss
22 m/d/yy h:mm
37 #,##0 ;(#,##0)
38 #,##0 ;[Red](#,##0)
39 #,##0.00;(#,##0.00)
40 #,##0.00;[Red](#,##0.00)
45 mm:ss
46 [h]:mm:ss
47 mmss.0
48 ##0.0E+0
49 @
-->

  <xsl:function name="xlsx2html:apply-format" as="xs:string">
    <xsl:param name="value" as="element(xls:v)"/>
    <xsl:param name="styleno" as="attribute(s)?"/>
    <xsl:variable name="numFmtId" as="xs:integer?" 
      select="root($value)//xls:cellXfs/xls:xf[xlsx2html:index-of(../*, .) = $styleno + 1]/@numFmtId"/>
    <xsl:variable name="formatCode" as="xs:string?" 
      select="root($value)//xls:numFmts/xls:numFmt[@numFmtId = $numFmtId]/@formatCode"/>
    <xsl:variable name="picture-string" as="xs:string?">
      <xsl:choose>
        <xsl:when test="$numFmtId = 14">
          <xsl:sequence select="'[Y]-[M,2]-[D,2]'"/>
        </xsl:when>
        <xsl:when test="matches(replace($formatCode, '\P{L}', ''), '^[ymd]+$')">
          <xsl:sequence select="
                                            replace(
                                              replace(
                                                replace(
                                                  replace(
                                                    replace(
                                                      replace(
                                                        replace(
                                                          replace(
                                                            replace(
                                                              $formatCode,
                                                              '(^\[\$-\d{3}\]|;@$)',
                                                              ''
                                                            ),
                                                            '\\(-)',
                                                            '$1'
                                                          ),
                                                          'mmmmm',
                                                          '[Mn,1-1]'
                                                        ),
                                                        'mmmm',
                                                        '[Mn]'
                                                      ),
                                                      'mmm',
                                                      '[Mn,3-3]'
                                                    ),
                                                    'mm',
                                                    '[M,2]'
                                                  ),
                                                  'm',
                                                  '[M]'
                                                ),
                                                'yyyy',
                                                '[Y]'
                                              ),
                                              'yy',
                                              '[Y,2]'
                                            )"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable> 
    <xsl:choose>
      <xsl:when test="$picture-string and normalize-space($value)">
        <xsl:variable name="day-string" select="concat('P', $value, 'D')" as="xs:string"/>
        <xsl:variable name="date" select="xs:date('1899-12-30') + xs:dayTimeDuration($day-string)" as="xs:date"/>
        <xsl:sequence select="format-date($date, $picture-string)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$value"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  
  
  <xsl:function name="xlsx2html:index-of" as="xs:integer*">
    <xsl:param name="nodes" as="node()*"/>
    <xsl:param name="node" as="node()*"/>
    <xsl:sequence select="index-of($nodes/generate-id(), $node/generate-id())"/>
  </xsl:function>

  <xsl:function name="xlsx2html:get-string-from-sst" as="node()*">
    <xsl:param name="index" as="xs:integer"/>
    <xsl:param name="sst" as="node()"/>
    <!-- <xsl:sequence select="$sst/xls:si[position() eq ($index+1)]//xls:t/node()"/> --><!--Include inline formatting (<r>, <rPr>) here! You may adapt this from wml2dbk, if possible.-->
    <xsl:apply-templates select="$sst/xls:si[position() eq ($index+1)]/node()" mode="props"/>
  </xsl:function>

  <xsl:template match="xls:f" mode="html">
    <xsl:param name="render" select="'no'"/>
    <xsl:if test="$render='yes'">
      <xsl:value-of select="."/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="xls:t" mode="html">
    <xsl:apply-templates select="node()" mode="#current"/>
  </xsl:template>


  <!-- mode: props -->

  <xsl:template match="xls:r" mode="props">
    <xsl:choose>
      <xsl:when test="xls:rPr">
        <xsl:element name="span">
          <xsl:apply-templates select="node()" mode="#current"/>
        </xsl:element>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="node()" mode="#current"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="xls:rPr" mode="props">
    <xsl:for-each select="*">
      <xsl:choose>
        <xsl:when test="$propmap//*[@name eq current()/name()]">
          <xsl:variable name="propmap-elem" select="$propmap//*[@name eq current()/name()]" as="node()"/>
          <xsl:attribute name="{$propmap-elem/@target-name}">
            <xsl:choose>
              <xsl:when test="$propmap-elem/@type='docx-boolean-prop'">
                <xsl:value-of select="$propmap-elem/@active"/>
              </xsl:when>
              <xsl:when test="$propmap-elem/@type=('docx-font-size', 'docx-font-family')"><!--$$$ conflicts with family-element and should be merged $$$-->
                <xsl:value-of select="@val"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:message select="concat('xlsx2html Error: Formatting property not handled correctly: ', local-name())"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="local-name()='family'">
          <xsl:attribute name="css:font-family">
            <xsl:choose><!-- Is this correct? -->
              <xsl:when test="@val='1'">serif</xsl:when><!--roman-->
              <xsl:when test="@val='2'">sans-serif</xsl:when><!--swiss-->
              <xsl:when test="@val='3'">monospace</xsl:when><!--modern-->
              <xsl:when test="@val='4'">cursive</xsl:when><!--script-->
              <xsl:when test="@val='5'">fantasy</xsl:when><!--decorative-->
            </xsl:choose>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:message select="concat('xlsx2html Warning: Formatting property could not be mapped: ', local-name())"/>
        </xsl:otherwise>        
      </xsl:choose>
      <xsl:apply-templates select="node()" mode="#current"/>
    </xsl:for-each>
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