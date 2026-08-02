<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_nullsafe_operator_in_ternary
// origin: languages/php/tests/php/test_php_expressions_ternary_coalescing_match.rs
// vybe-test-mode: compile

class Profile { public string $avatar = "avatar.png"; }
class User { public ?Profile $profile = null; }

$user = new User();
$avatar = $user?->profile ? $user->profile->avatar : "default.png";
echo $avatar;
