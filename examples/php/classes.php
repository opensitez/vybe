<?php

class Animal {
    public $name;
    public $sound;

    public function __construct($name, $sound) {
        $this->name = $name;
        $this->sound = $sound;
    }

    public function speak() {
        return $this->name . " says " . $this->sound;
    }
}

$dog = new Animal("Dog", "Woof");
$cat = new Animal("Cat", "Meow");
echo $dog->speak();
echo $cat->speak();

// Inheritance
class Dog extends Animal {
    public $breed;

    public function __construct($name, $breed) {
        parent::__construct($name, "Woof");
        $this->breed = $breed;
    }

    public function info() {
        return $this->name . " (" . $this->breed . ")";
    }
}

$rex = new Dog("Rex", "German Shepherd");
echo $rex->speak();
echo $rex->info();

// Static methods
class MathHelper {
    public static function add($a, $b) {
        return $a + $b;
    }

    public static function max($a, $b) {
        if ($a > $b) {
            return $a;
        }
        return $b;
    }
}

echo MathHelper::add(3, 4);
echo MathHelper::max(10, 20);

// Class with methods
class Counter {
    public $count;

    public function __construct() {
        $this->count = 0;
    }

    public function increment() {
        $this->count++;
    }

    public function decrement() {
        $this->count--;
    }

    public function value() {
        return $this->count;
    }
}

$c = new Counter();
$c->increment();
$c->increment();
$c->increment();
$c->decrement();
echo $c->value();
