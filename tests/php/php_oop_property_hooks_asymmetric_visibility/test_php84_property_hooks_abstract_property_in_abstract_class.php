<?php
// vybe-test: php/php_oop_property_hooks_asymmetric_visibility/test_php84_property_hooks_abstract_property_in_abstract_class
// origin: languages/php/tests/php/test_php_oop_property_hooks_asymmetric_visibility.rs
// vybe-test-mode: compile

abstract class Widget {
    abstract public string $label { get; }
}

class ButtonWidget extends Widget {
    public string $label { get => "Click Me"; }
}

$b = new ButtonWidget();
echo $b->label;
