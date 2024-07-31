using System.Collections.Generic;

namespace CongratulatoryEmailWorkflow
{
    public static class EmailConstants
    {
        public static Dictionary<string, string> EmailTemplates = new Dictionary<string, string>()
        {
            {
                "Birthday Congratulation",
                @"<?xml version=""1.0"" ?><xsl:stylesheet xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" version=""1.0""><xsl:output method=""text"" indent=""no""/><xsl:template match=""/data""><![CDATA[<div>Alles Gute zum Geburtstag, [GendercodeTitle] [Firstname] [Lastname]!</div><div class=keyboardFocusClass><div><br></div><div>Heute ist [Birthdate]. Ich hoffe, diese Nachricht findet dich gut und du hast eine wundervolle Geburtstagsfeier. Möge dein Tag voller Freude, Lachen und der Gesellschaft derer sein, die dir lieb und teuer sind.</div><div><br></div><div>Ich wünsche Ihnen alles Gute an Ihrem besonderen Tag und für das kommende Jahr. Genieße deinen Geburtstag, [Firstname] [Lastname]!</div></div>]]></xsl:template></xsl:stylesheet>"
            }
        };
    }
}
