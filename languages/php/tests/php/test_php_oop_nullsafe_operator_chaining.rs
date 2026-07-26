use super::helpers::run_prints;

#[test]
fn test_nullsafe_operator_chain_on_null() {
    assert_eq!(
        run_prints(
            r#"<?php
class Profile {
    public ?Address $address = null;
}
class Address {
    public function getCity(): string { return 'London'; }
}
$p = new Profile();
echo ($p->address?->getCity() ?? 'no_city'), "\n";
"#
        ),
        vec!["no_city"]
    );
}

#[test]
fn test_nullsafe_operator_chain_on_non_null() {
    assert_eq!(
        run_prints(
            r#"<?php
class UserProfile {
    public ?AddressInfo $address = null;
}
class AddressInfo {
    public function getCity(): string { return 'Paris'; }
}
$p = new UserProfile();
$p->address = new AddressInfo();
echo ($p->address?->getCity() ?? 'no_city'), "\n";
"#
        ),
        vec!["Paris"]
    );
}

#[test]
fn test_nullsafe_property_access_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
class User {
    public function __construct(public ?Profile $profile = null) {}
}
class Profile {
    public function __construct(public ?AddressInfo $address = null) {}
}
class AddressInfo {
    public function __construct(public string $city = 'Rome') {}
}
$userWithCity = new User(new Profile(new AddressInfo('Milan')));
$userWithout = new User();
echo $userWithCity->profile?->address?->city ?? 'none';
echo '|';
echo $userWithout->profile?->address?->city ?? 'none';
"#,
        ),
        vec!["Milan|none"]
    );
}

#[test]
fn test_nullsafe_method_then_property_and_arithmetic() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    public function value(): int { return 4; }
}
class Holder {
    public ?Counter $counter;
    public function __construct(?Counter $counter = null) { $this->counter = $counter; }
}
$has = new Holder(new Counter());
$none = new Holder();
echo ($has->counter?->value() + 1);
echo '|';
echo ($none->counter?->value() + 1) ?? 'none';
"#
        ),
        vec!["5|none"]
    );
}

#[test]
fn test_nullsafe_dynamic_method_name_precedence() {
    assert_eq!(
        run_prints(
            r#"<?php
class Service {
    public ?Tag $ghost;
    public function tag(): ?Tag {
        return new Tag();
    }
    public function fallback(): string {
        return 'fb';
    }
    public function __construct() {
        $this->ghost = null;
    }
}
class Tag {
    public function name(): string { return 'ok'; }
}

$service = new Service();
echo $service->tag()?->name() . '|';
echo $service->ghost?->name() ?? 'missing';
"#
        ),
        vec!["ok|missing"]
    );
}

#[test]
fn test_nullsafe_with_ternary_and_space_operator_style() {
    assert_eq!(
        run_prints(
            r#"<?php
class Score {
    public int $value;
    public function __construct(int $value) { $this->value = $value; }
}
class Box {
    public ?Score $score;
    public function __construct(?Score $score = null) { $this->score = $score; }
}

$with = new Box(new Score(7));
$without = new Box(null);
$left = $with->score?->value;
$right = $without->score?->value;
echo (($left <=> 5) > 0) ? 'gt' : 'lte';
echo '|';
echo (($right <=> 5) > 0) ? 'gt' : (($right ?? 0) ? 'truthy' : 'false');
"#
        ),
        &["gt|false"]
    );
}

#[test]
fn test_nullsafe_operator_truthiness_in_conditions() {
    assert_eq!(
        run_prints(
            r#"<?php
class Node {
    public function score(): int { return 0; }
}
class Container {
    public ?Node $node = null;
}

$ready = new Container();
$ready->node = new Node();
echo $ready->node?->score() ? 'truthy' : 'falsey';
echo '|';
$empty = new Container();
echo $empty->node?->score() ? 'truthy' : 'falsey';
echo '|';
echo (($empty->node?->score() ?: 'fallback'));
"#
        ),
        vec!["falsey|falsey|fallback"]
    );
}

#[test]
fn test_nullsafe_in_call_chain_with_coalesce_precedence() {
    assert_eq!(
        run_prints(
            r#"<?php
class Level {
    public function name(): string { return ''; }
}
class Holder {
    public ?Level $level = null;
}
class Root {
    public function holder(): ?Holder { return null; }
    public function fallback(): string { return 'fb'; }
}
class RootWithLevel extends Root {
    public Holder $holderObj;
    public function __construct() {
        $this->holderObj = new Holder();
        $this->holderObj->level = new Level();
    }
    public function holder(): ?Holder { return $this->holderObj; }
}

echo (new Root())->holder()?->name() ?? 'no-holder';
echo '|';
echo (new RootWithLevel())->holder()?->level?->name() ?? 'no-name';
echo '|';
echo (new RootWithLevel())->holder()?->level?->name() ?: 'fallback-name';
"#
        ),
        vec!["no-holder|no-name|fallback-name"]
    );
}

#[test]
fn test_nullsafe_with_property_read_after_method_call() {
    assert_eq!(
        run_prints(
            r#"<?php
class Child {
    public function __construct(public ?Grandchild $nested = null) {}
    public function child(): ?Grandchild { return $this->nested; }
}
class Grandchild {
    public string $value = 'nested';
}
class Parent {
    public Child $child;
    public function __construct(?Child $child = null) { $this->child = $child ?? new Child(); }
}

echo (new Parent(new Child(new Grandchild())))->child()->child?->value ?? 'none';
echo '|';
echo (new Parent(new Child()))->child()->child?->value ?? 'none';
"#
        ),
        vec!["nested|none"]
    );
}
