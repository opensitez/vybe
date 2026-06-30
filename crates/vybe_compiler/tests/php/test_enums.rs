//! Unit and backed `enum` declarations, methods, and `tryFrom`.

crate::php_cases! {
    backed_enum_value_property => {
        r#"<?php
enum HttpStatus: int { case Ok = 200; case NotFound = 404; }
echo HttpStatus::NotFound->value;
"#,
        ["404"]
    };

    pure_enum_case_name => {
        r#"<?php
enum Direction { case North; case South; }
echo Direction::North->name;
"#,
        ["North"]
    };

    enum_try_from_returns_case => {
        r#"<?php
enum Role: string { case Admin = 'admin'; case Guest = 'guest'; }
echo Role::tryFrom('admin')?->name ?? 'none';
"#,
        ["Admin"]
    };

    enum_try_from_invalid_returns_null => {
        r#"<?php
enum Role: string { case Admin = 'admin'; }
echo Role::tryFrom('missing') === null ? 'null' : 'found';
"#,
        ["null"]
    };

    enum_from_throws_on_invalid_value => {
        r#"<?php
enum Bit: int { case Zero = 0; case One = 1; }
try { Bit::from(9); echo 'ok'; } catch (ValueError) { echo 'bad'; }
"#,
        ["bad"]
    };

    enum_cases_lists_all_members => {
        r#"<?php
enum Size { case S; case M; case L; }
echo count(Size::cases());
"#,
        ["3"]
    };

    enum_method_on_backed_enum => {
        r#"<?php
enum Priority: int {
    case Low = 1;
    case High = 3;
    public function label(): string {
        return match ($this) { self::Low => 'low', self::High => 'high' };
    }
}
echo Priority::High->label();
"#,
        ["high"]
    };

    enum_implements_interface_method => {
        r#"<?php
interface Labeled { public function label(): string; }
enum Color: string implements Labeled {
    case Red = 'r';
    public function label(): string { return 'red'; }
}
echo (Color::Red)->label();
"#,
        ["red"]
    };

    enum_comparison_same_case_is_identical => {
        r#"<?php
enum Mode { case On; case Off; }
echo Mode::On === Mode::On ? 'same' : 'diff';
"#,
        ["same"]
    };

    enum_in_array_keyed_by_name => {
        r#"<?php
enum Tier: string { case Free = 'free'; case Pro = 'pro'; }
$map = [Tier::Free->name => 0, Tier::Pro->name => 1];
echo $map['Pro'];
"#,
        ["1"]
    };

    enum_switch_via_match => {
        r#"<?php
enum State { case Open; case Closed; }
function code(State $s): string {
    return match ($s) { State::Open => 'O', State::Closed => 'C' };
}
echo code(State::Closed);
"#,
        ["C"]
    };

    backed_enum_string_cast_via_value => {
        r#"<?php
enum Code: string { case A = 'alpha'; }
echo Code::A->value;
"#,
        ["alpha"]
    };
}
