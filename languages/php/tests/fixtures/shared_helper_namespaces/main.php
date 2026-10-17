<?php
namespace HelperConsumer;

use HelperFirst\Same as FirstSame;
use HelperSecond\Same as SecondSame;
use function HelperFirst\label as first_label;
use function HelperSecond\label as second_label;

require_once __DIR__ . '/first.inc';
include __DIR__ . '/second.inc';

$first = new FirstSame();
$second = new SecondSame();
return first_label(7) . ':' . $first->value() . '|' .
    second_label(7) . ':' . $second->value() . '|' . __NAMESPACE__;
