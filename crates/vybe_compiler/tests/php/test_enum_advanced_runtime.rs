//! Backed enums — `from`, serialization, traits, and match (beyond `test_enums.rs`).

crate::php_cases! {
    backed_string_enum_value_roundtrip => {
        r#"<?php
enum Status: string { case Ok = 'ok'; case Err = 'err'; }
echo Status::from('ok')->value;
"#,
        ["ok"]
    };

    backed_int_enum_increment_value => {
        r#"<?php
enum Level: int { case Low = 1; case High = 3; }
echo Level::High->value - Level::Low->value;
"#,
        ["2"]
    };

    enum_cases_names_array => {
        r#"<?php
enum Color: string { case Red = 'r'; case Blue = 'b'; }
echo implode(',', array_map(fn($c) => $c->name, Color::cases()));
"#,
        ["Red,Blue"]
    };

    enum_in_array_unique => {
        r#"<?php
enum Mode { case On; case Off; }
$a = [Mode::On, Mode::On, Mode::Off];
echo count(array_unique($a, SORT_REGULAR));
"#,
        ["2"]
    };

    enum_match_exhaustive => {
        r#"<?php
enum Dir { case N; case S; }
function flip(Dir $d): string { return match ($d) { Dir::N => 'north', Dir::S => 'south' }; }
echo flip(Dir::N);
"#,
        ["north"]
    };

    enum_method_static_cases_count => {
        r#"<?php
enum Tier: int { case Free = 0; case Pro = 1; case Ent = 2; }
echo count(Tier::cases());
"#,
        ["3"]
    };

    enum_implements_json_serializable => {
        r#"<?php
enum Code: int implements JsonSerializable {
    case A = 1;
    public function jsonSerialize(): int { return $this->value; }
}
echo json_encode(Code::A);
"#,
        ["1"]
    };

    enum_unit_serialize_name => {
        r#"<?php
enum Flag { case Alpha; case Beta; }
echo Flag::Alpha->name;
"#,
        ["Alpha"]
    };

    enum_backed_try_from_int => {
        r#"<?php
enum Http: int { case Ok = 200; case Teapot = 418; }
echo Http::tryFrom(418)?->name ?? 'x';
"#,
        ["Teapot"]
    };

    enum_comparison_different_cases => {
        r#"<?php
enum Bit { case Zero; case One; }
echo Bit::Zero === Bit::One ? 'same' : 'diff';
"#,
        ["diff"]
    };

    enum_property_on_class => {
        r#"<?php
enum State { case Active; case Idle; }
class Machine { public State $s = State::Idle; }
echo (new Machine())->s->name;
"#,
        ["Idle"]
    };

    enum_switch_case => {
        r#"<?php
enum Pet { case Cat; case Dog; }
$p = Pet::Cat;
echo match ($p) { Pet::Cat => 'meow', Pet::Dog => 'woof' };
"#,
        ["meow"]
    };

    enum_backed_in_match_value => {
        r#"<?php
enum Num: int { case One = 1; case Two = 2; }
$n = Num::Two;
echo $n->value;
"#,
        ["2"]
    };

    enum_trait_shared_method => {
        r#"<?php
trait Named { public function slug(): string { return strtolower($this->name); } }
enum Planet { use Named; case Mars; }
echo Planet::Mars->slug();
"#,
        ["mars"]
    };

    enum_from_array_map => {
        r#"<?php
enum N: int { case A = 1; case B = 2; }
echo implode(',', array_map(fn(N $n) => $n->value, N::cases()));
"#,
        ["1,2"]
    };

    enum_readonly_class_field => {
        r#"<?php
enum Role: string { case Admin = 'admin'; }
readonly class User { public function __construct(public Role $role) {} }
echo (new User(Role::Admin))->role->value;
"#,
        ["admin"]
    };

    enum_generator_yield_cases => {
        r#"<?php
enum N { case X; case Y; }
function all(): Generator { foreach (N::cases() as $c) { yield $c->name; } }
echo implode('', iterator_to_array(all()));
"#,
        ["XY"]
    };

    enum_backed_string_in_string_context => {
        r#"<?php
enum Lang: string { case En = 'en'; }
echo 'lang=' . Lang::En->value;
"#,
        ["lang=en"]
    };

    enum_unit_in_array_keys_not_allowed_use_list => {
        r#"<?php
enum K { case A; case B; }
$list = [K::A, K::B];
echo $list[1]->name;
"#,
        ["B"]
    };

    enum_interface_default_method => {
        r#"<?php
interface Valued { public function v(): int; }
enum Score: int implements Valued { case Low = 1; public function v(): int { return $this->value; } }
echo Score::Low->v();
"#,
        ["1"]
    };

    enum_attributes_on_case => {
        r#"<?php
#[\AllowDynamicProperties]
enum E { case One; }
echo E::One->name;
"#,
        ["One"]
    };

    enum_private_method_via_public => {
        r#"<?php
enum E {
    case A;
    private function secret(): string { return 's'; }
    public function reveal(): string { return $this->secret(); }
}
echo E::A->reveal();
"#,
        ["s"]
    };

    enum_backed_from_valid => {
        r#"<?php
enum Port: int { case Http = 80; case Https = 443; }
echo Port::from(443)->name;
"#,
        ["Https"]
    };

    enum_foreach_cases => {
        r#"<?php
enum Axis { case X; case Y; case Z; }
$c = 0;
foreach (Axis::cases() as $_) { $c++; }
echo $c;
"#,
        ["3"]
    };

    enum_in_union_return => {
        r#"<?php
enum Flag { case On; case Off; }
function f(): Flag|string { return Flag::On; }
echo f()->name;
"#,
        ["On"]
    };
}
