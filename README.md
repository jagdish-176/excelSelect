<p>excelSelect<br />
===========<br />
Simple jQuery plugin to give you excel like selection on &#39;table&#39; tag.<br />
1.Shift/Ctrl cell selection supported.<br />
2. rowspan , colspan handled.</p>

<p>Usage :<br />
==========<br />
include<br />
&lt;link href=&quot;excel-select.css&quot; media=&quot;all&quot; rel=&quot;stylesheet&quot; type=&quot;text/css&quot;&gt;<br />
&lt;script type=&quot;text/javascript&quot; src=&quot;excel-select.js&quot;&gt;&lt;/script&gt;<br />
in your html .<br />
and get ready for ..<br />
$(document).ready(function() {<br />
$(&#39;table&#39;).excelSelect({<br />
&nbsp;&nbsp; &nbsp;onSelectionEnd:function($table,api){<br />
&nbsp;&nbsp; &nbsp;&nbsp;&nbsp; &nbsp;console.log(&quot;Selected cell [&quot;+api.getSelectedCells().length+&quot;]&quot;);<br />
&nbsp;&nbsp; &nbsp;}});<br />
});</p>
