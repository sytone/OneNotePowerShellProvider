<?xml version="1.0" encoding="utf-8"?>

<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

<xsl:template match="OneNoteDigest">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Web Notebooks Digest</title>
<style type="text/css">
    th {
    font-family: Calibri, Arial, sans-serif;
    background-color: #17365D;
    color: #FFFFFF;
    }
    body {
    font-family: Calibri, Arial, sans-serif;
    }
    #masthead {
    background-color: #325187;
    font-family: Segoe UI, Verdana, Arial, Sans-Serif;
    padding:20px;
    }
    #logotype{
    background-color: #325187;
    font-family: Segoe UI, Verdana, Arial, Sans-Serif;
    padding:20px;
    }

    #logotype .logotitle {
    color:#FFFFFF;
    font-size:100%;
    display:block;
    }

    #logotype .logosubtitle {
    color:#f46d0b;
    letter-spacing:3px;
    display:block;
    font-size:90%;
    }
    .publishedDate{
    top:0px;
    font-size:75%;
    color:#FFFFFF;
    }


    #pageTitleH1
    {
    color:#ffffff;
    font-family: Segoe UI, Verdana, Arial, Sans-Serif;
    font-weight: normal;
    font-size: 110%;
    }
    a {
    text-decoration:none;
    }
    a:hover {
    text-decoration:underline;
    }

</style>
</head>

<body>
    <table cellspacing="0" width="100%">
        <col width="200px"/>
        <tr>
            <td id="logotype">
                <a class="logotitle" href="http://officelabs">officelabs</a>
                <div class="logosubtitle">web notebooks</div>
            </td>
            <td id="masthead">
                <div id="pageTitleH1" align="left">
                    <xsl:value-of select="NotebookName"/> Digest
                </div>
                <div class="publishedDate">
                    Changes since <xsl:value-of select="ChangesSince"/>
                </div>
            </td>
        </tr>
    </table>
    <p>
        This is an automatic mail message that contains the pages in the <xsl:value-of select="NotebookName"/> notebook
        that have changed since <xsl:value-of select="ChangesSince" />.
        You will find each changed page as an attachment to this message.
    </p>
    <table>
        <tr>
            <th>Page</th>
            <th>Link</th>
        </tr>
        <xsl:apply-templates select="ChangedPage" />
    </table>
</body>
</html>

</xsl:template>

    <xsl:template match="ChangedPage">
        <tr xmlns="http://www.w3.org/1999/xhtml">
            <td>
                <xsl:value-of select="Name"/>
            </td>
            <td>
                <a>
                    <xsl:attribute name="href">
                        <xsl:value-of select="Hyperlink"/>
                    </xsl:attribute>
                    Open in OneNote
                </a>
            </td>
        </tr>
    </xsl:template>
</xsl:stylesheet> 
