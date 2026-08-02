<?php
// vybe-test: php/reflection/reflection_method_name_and_visibility
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class Service {
    public function doWork(): void {}
    protected function helper(): void {}
    private function internal(): void {}
}
$rc = new ReflectionClass(Service::class);
$method = $rc->getMethod('doWork');
echo $method->getName();
echo $method->isPublic() ? ':public' : ':not-public';
