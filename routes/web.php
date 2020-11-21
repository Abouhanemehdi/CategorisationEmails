<?php

use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

// Route::get('/', function () {return view('mailbox');});

Route::get('/', 'OutlookMailBoxController@login');
Route::get('/login', 'OutlookMailBoxController@login');
Route::get('/mailbox', 'OutlookMailBoxController@show');
Route::get('/csvexport', 'OutlookMailBoxController@csvExport');
Route::get('/signin', 'OutlookAuthController@signin');
Route::get('/callback', 'OutlookAuthController@callback');
Route::get('/signout', 'OutlookAuthController@signout');
