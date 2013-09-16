<html>sample code.
<?php

class IEEventSinker {
  var $completed = false;
  var $url = '';

  function DocumentComplete(&$dom, $url) {
    if ($url == $this->url) {
      $this->completed = true;
    }
  }

  function OnQuit() {
    $this->completed = false;
  }
}

$com_ie = new COM('InternetExplorer.Application', null, CP_UTF8);
$com_ie->visible = true;

$sink = new IEEventSinker();
com_event_sink($com_ie, $sink, 'DWebBrowserEvents2');

$sink->url = 'http://search.yahoo.co.jp/';
$com_ie->navigate($sink->url);

$st = time();
while (!$sink->completed) {
  com_message_pump(3000);
  if ((time() - $st) >= 5) break;
}

$com_ie->document->forms(0)->p->value = 'php';
$com_ie->document->forms(0)->submit(null);
