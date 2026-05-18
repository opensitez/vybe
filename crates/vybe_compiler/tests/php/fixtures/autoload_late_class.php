<?php
echo 'autoload';

class LateLoadedClass {
    public function __construct() {
        echo 'ctor';
    }
}