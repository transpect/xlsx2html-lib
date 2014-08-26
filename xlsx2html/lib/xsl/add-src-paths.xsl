<xsl:stylesheet 
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
    xmlns:functx="http://www.functx.com"
    version="2.0">
  
  <xsl:import href="http://www.functx.com/XML_Elements_and_Attributes/XML_Document_Structure/path-to-node-with-pos.xsl"/>
  
  <xsl:template match="@*|node()" priority="-10">
    <xsl:copy>
      <xsl:apply-templates select="@*|node()"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="*">
    <xsl:copy>
      <xsl:attribute name="srcpath" select="concat(/*/@xml:base, '?xpath=/', functx:path-to-node-with-pos(.))"/>
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>

</xsl:stylesheet>
