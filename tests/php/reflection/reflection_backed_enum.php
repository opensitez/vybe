<?php
// vybe-test: php/reflection/reflection_backed_enum
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

enum Status: string { case Active = 'active'; case Inactive = 'inactive'; }
$re = new ReflectionEnum(Status::class);
echo $re->isBacked() ? 'backed' : 'pure';
echo ':' . $re->getBackingType()->getName();
