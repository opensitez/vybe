<?php
// vybe-test: php/namespaces/namespace_enum
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Domain\Status;
enum OrderStatus: string {
    case Pending  = 'pending';
    case Shipped  = 'shipped';
    case Delivered = 'delivered';
}

namespace App;
use Domain\Status\OrderStatus;
$status = OrderStatus::Shipped;
echo $status->value;
