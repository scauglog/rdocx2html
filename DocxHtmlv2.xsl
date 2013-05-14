<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:WX="http://schemas.microsoft.com/office/word/2003/auxHint"
    xmlns:aml="http://schemas.microsoft.com/aml/2001/core"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    version="1.0">
	
	<xsl:template match="/">
		<html>
			<body>
				<p>XSL test</p>
				<xsl:apply-templates/>
			</body>
		</html>
	</xsl:template>

	<xsl:template match="w:commentRangeStart"><xsl:text disable-output-escaping="yes">&lt;em&gt;</xsl:text></xsl:template>
	<xsl:template match="w:commentRangeEnd"><xsl:text disable-output-escaping="yes">&lt;/em&gt;</xsl:text></xsl:template>

	<xsl:template match="w:p">
		<xsl:variable name="style" select="w:pPr/w:pStyle/@w:val"/>
		<xsl:apply-templates select="w:commentRangeStart"/>

		<xsl:choose>
			<xsl:when test="$style='Titre1'">
				<h1>
					<xsl:for-each select="w:r">
						<xsl:value-of select="w:t"/>
					</xsl:for-each>
				</h1>
			</xsl:when>
			<xsl:when test="$style='Titre2'">
				<h2>
					<xsl:apply-templates select="w:r"/>
				</h2>
			</xsl:when>
			<xsl:when test="$style='Titre3'">
				<h3>
					<xsl:apply-templates select="w:r"/>
				</h3>
			</xsl:when>
			<xsl:when test="$style='Titre4'">
				<h4>
					<xsl:apply-templates select="w:r"/>
				</h4>
			</xsl:when>
			<xsl:when test="$style='Titre5'">
				<h5>
					<xsl:apply-templates select="w:r"/>
				</h5>
			</xsl:when>
			<xsl:otherwise>
				<p>
					<xsl:apply-templates select="w:r"/>
				</p>
			</xsl:otherwise>
		
		</xsl:choose>
		<xsl:apply-templates select="w:commentRangeEnd"/>
	</xsl:template>
	
	<xsl:template match="w:r">
		<xsl:value-of select="w:t"/>
	</xsl:template>
	
	
	<xsl:template match="w:tbl">
		<table>
			<xsl:apply-templates select="w:tr"/>
		</table>
	</xsl:template>
	
	<xsl:template match="w:tr">
		<tr>
			<xsl:apply-templates select="w:tc"/>
		</tr>
	</xsl:template>
	
	<xsl:template match="w:tc">
		<td>
			<xsl:apply-templates select="w:p"/>
		</td>
	</xsl:template>
</xsl:stylesheet>