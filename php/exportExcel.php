<?php
require 'vendor/autoload.php';
ini_set('memory_limit', '1024G');
use Medoo\Medoo;

$database = new Medoo([
    'database_type' => 'mysql',
    'database_name' => 'test',
    'server' => '127.0.0.1',
    'username' => 'root',
    'password' => 'root'
]);

// Enjoy
$database->insert('account', [
    'user_name' => 'foo',
    'email' => 'foo@bar.com'
]);

$data = $database->select("test_data", '*');

echo json_encode($data);
