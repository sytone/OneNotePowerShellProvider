<?xml version="1.0" encoding="utf-8"?>

<xsl:stylesheet version="1.0"
    xmlns:one="http://schemas.microsoft.com/office/onenote/2007/onenote"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

    <xsl:output method="text" />

<xsl:template match="/">
    <xsl:apply-templates select="//one:Outline" />
</xsl:template>

    <xsl:template match="one:Outline">
        <xsl:apply-templates select="*//one:T" />
    </xsl:template>

    <xsl:template match="one:T">
        <xsl:value-of select="." disable-output-escaping="yes"/>
        <xsl:text>
</xsl:text>
    </xsl:template>

</xsl:stylesheet> 
