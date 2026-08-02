<?php
// vybe-test: php/php84/property_hook_computed
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

class Circle {
    public float $radius = 0.0;
    public float $area {
        get { return M_PI * $this->radius ** 2; }
    }
    public float $circumference {
        get { return 2 * M_PI * $this->radius; }
    }
}
$c = new Circle();
$c->radius = 5.0;
echo round($c->area, 4) . ':' . round($c->circumference, 4);
