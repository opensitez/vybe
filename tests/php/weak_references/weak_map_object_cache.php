<?php
// vybe-test: php/weak_references/weak_map_object_cache
// origin: languages/php/tests/php/test_weak_references.rs
// vybe-test-mode: compile

class User { public function __construct(public int $id, public string $name) {} }
$computed = new WeakMap();
function getDisplayName(User $user, WeakMap $cache): string {
    if (!isset($cache[$user])) {
        $cache[$user] = strtoupper($user->name) . '#' . $user->id;
    }
    return $cache[$user];
}
$u = new User(1, 'alice');
echo getDisplayName($u, $computed);
echo getDisplayName($u, $computed);
