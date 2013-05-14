<?xml version="1.0" encoding="utf-8"?>

<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:WX="http://schemas.microsoft.com/office/word/2003/auxHint"
    xmlns:aml="http://schemas.microsoft.com/aml/2001/core"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
	xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    version="1.0">
	<xsl:output method="html" omit-xml-declaration="no" indent="yes"/>
	<xsl:param name="h1" select="'Titre1'" />
	<xsl:param name="h2" select="'Titre2'" />
	<xsl:param name="h3" select="'Titre3'" />
	<xsl:param name="h4" select="'Titre4'" />
	<xsl:param name="h5" select="'Titre5'" />
	<xsl:param name="list" select="'Paragraphedeliste'" />
	<xsl:template match="/">
		<html>
			<body>
				<div class="uncommented">
					<xsl:apply-templates/>
				</div>
			</body>
		</html>
	</xsl:template>

	<xsl:template match="w:commentRangeStart"><xsl:text disable-output-escaping="yes">&lt;/div&gt; &lt;div class="commented"&gt;</xsl:text></xsl:template>
	<xsl:template match="w:commentRangeEnd"><xsl:text disable-output-escaping="yes">&lt;/div&gt; &lt;div class="uncommented"&gt;</xsl:text></xsl:template>

	<xsl:template match="w:p">
		<xsl:variable name="style" select="w:pPr/w:pStyle/@w:val"/>
		<xsl:apply-templates select="w:commentRangeStart"/>

		<xsl:choose>
			<xsl:when test="$style=$h1">
				<h1>
					<xsl:for-each select="w:r">
						<xsl:value-of select="w:t"/>
					</xsl:for-each>
				</h1>
			</xsl:when>
			<xsl:when test="$style=$h2">
				<h2>
					<xsl:apply-templates select="w:r"/>
				</h2>
			</xsl:when>
			<xsl:when test="$style=$h3">
				<h3>
					<xsl:apply-templates select="w:r"/>
				</h3>
			</xsl:when>
			<xsl:when test="$style=$h4">
				<h4>
					<xsl:apply-templates select="w:r"/>
				</h4>
			</xsl:when>
			<xsl:when test="$style=$h5">
				<h5>
					<xsl:apply-templates select="w:r"/>
				</h5>
			</xsl:when>
			<xsl:when test="$style=$list">
				<ul>
					<li>
						<xsl:apply-templates select="w:r"/>
					</li>
				</ul>
			
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
		<xsl:apply-templates select="w:drawing"/>
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
	
	
	<xsl:template match="w:drawing">
		<xsl:apply-templates select="descendant::a:blip" />
	</xsl:template>

	<xsl:template match="a:blip">
		<xsl:variable name="relid" select="@r:embed"/>
		<xsl:element name="img">
				<xsl:attribute name="alt">
					<xsl:value-of select="$relid"/>
				</xsl:attribute>
				<xsl:attribute name="src">
					<xsl:value-of select="document('./word/_rels/document.xml.rel')/rel:Relationships/rel:Relationship[@Id=$relid]/@Target"/>
				</xsl:attribute>
		</xsl:element>
	</xsl:template>
	
</xsl:stylesheet>