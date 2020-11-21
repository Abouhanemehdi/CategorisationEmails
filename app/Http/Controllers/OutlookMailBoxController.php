<?php

namespace App\Http\Controllers;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use Microsoft\Graph\Graph;
use Microsoft\Graph\Model;
use App\TokenStore\TokenCache;
use App\TimeZones\TimeZones;

class OutlookMailBoxController extends Controller
{
  public function show()
  {
    $viewData = $this->loadViewData();
    $graph = $this->getGraph();
    $queryParams = array(
      '$select' => 'subject,sender, receivedDateTime, hasAttachments',
      '$top' => 25
    );

    // Append query parameters to the '/me/calendarView' url
    $getMsgsUrl = '/me/messages?'.http_build_query($queryParams);

    $ptr = $graph->createCollectionRequest('GET', $getMsgsUrl)
      // Add the user's timezone to the Prefer header
      ->setReturnType(Model\Message::class);

    $msgs= $ptr->getPage();

    $viewData['msgs'] = $msgs;
    return view('mailbox', $viewData);

    // return response()->json($msgs);
  }

  private function getGraph(): Graph
  {
    // Get the access token from the cache
    $tokenCache = new TokenCache();
    $accessToken = $tokenCache->getAccessToken();

    // Create a Graph client
    $graph = new Graph();
    $graph->setAccessToken($accessToken);
    return $graph;
  }


  public function csvExport(Request $request)
  {
    $viewData = $this->loadViewData();
    $graph = $this->getGraph();
    $queryParams = array(
      '$select' => 'subject,sender, receivedDateTime, hasAttachments',
      'top' => 200

    );

    // Append query parameters to the '/me/calendarView' url
    $getMsgsUrl = '/me/messages?'.http_build_query($queryParams);

    $msgs = $graph->createRequest('GET', $getMsgsUrl)
      // Add the user's timezone to the Prefer header
      ->setReturnType(Model\Message::class)
      ->execute();
  
    
    $headers = array(
      "Content-Encoding" => "UTF-8",
      "Content-type" => "text/csv; charset=UTF-8",
      "Content-Disposition" => "attachment; filename=file.csv",
      "Pragma" => "no-cache",
      "Cache-Control" => "must-revalidate, post-check=0, pre-check=0",
      "Expires" => "0"
    );

    
    $columns = array('emailID', 'receivedDateTime', 'hasAttachments', 'senderName', 'senderEmail', 'subject', 'body');

    $callback = function() use ($msgs, $columns)
    {
        $file = fopen('php://output', 'w');
        fputcsv($file, $columns);

        foreach($msgs as $msg) {
            $id= $msg->getId();
            $receivedDateTime= \Carbon\Carbon::parse($msg->getReceivedDateTime())->format('j/n g:i A');
            $hasAttachments= $msg->getHasAttachments();
            $senderName= $msg->getSender()->getEmailAddress()->getName();
            $senderAdresse= $msg->getSender()->getEmailAddress()->getAddress();
            $subject= $msg->getSubject();
            $body= "hello word";
            fputcsv($file, array($id, $receivedDateTime, $hasAttachments, $senderName, $senderAdresse, $subject, $body ));
        }
        fclose($file);
    };
    
    return response()->stream($callback, 200, $headers);
  }

  public function login()
  {
    return view('login');
  }
}









