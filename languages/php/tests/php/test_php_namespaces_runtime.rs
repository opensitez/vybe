use super::helpers::run_prints;

fn assert_int(source: &str, expected: i64) {
    assert_eq!(run_prints(source), vec![expected.to_string()]);
}

fn assert_str(source: &str, expected: &str) {
    assert_eq!(run_prints(source), vec![expected.to_string()]);
}

#[test]
fn php_namespace_runtime() {
    let top_levels = [
        "App", "Domain", "Platform", "Modules", "Service", "Core", "Shared", "Http", "Cli",
        "Storage", "Worker",
    ];
    let feature_groups = [
        "Auth",
        "Catalog",
        "Orders",
        "Billing",
        "Notifications",
        "Search",
        "Reports",
        "Analytics",
        "Scheduler",
        "Routing",
    ];

    for top_index in 0..top_levels.len() {
        for feature_index in 0..feature_groups.len() {
            let top = top_levels[top_index];
            let feature = feature_groups[feature_index];
            let index = (top_index * 100 + feature_index + 1) as i64;

            let ns = format!("{}\\\\{}", top, feature);

            let class_name = format!("Svc{top_index}_{feature_index}");
            let class_alias = format!("Alias{top_index}_{feature_index}");
            let class_src = format!(
                "<?php\nnamespace {ns};\nclass {class_name} {{\n    public function id(): int {{\n        return {index};\n    }}\n}}\n\nnamespace App;\nuse {ns}\\{class_name} as {class_alias};\n\necho (new {class_alias}())->id();\n",
            );
            assert_int(&class_src, index);

            let fn_name = format!("value_for_{top_index}_{feature_index}");
            let fn_src = format!(
                "<?php\nnamespace {ns};\nfunction {fn_name}(): int {{\n    return {index};\n}}\n\nnamespace Workspace;\nuse function {ns}\\{fn_name};\necho {fn_name}();\n",
            );
            assert_int(&fn_src, index);

            let const_name = format!("SERVICE_ID_{top_index}_{feature_index}");
            let const_src = format!(
                "<?php\nnamespace {ns};\nconst {const_name} = {index};\n\nnamespace Workspace;\nuse const {ns}\\{const_name};\necho {const_name};\n",
            );
            assert_int(&const_src, index);
        }
    }
}

#[test]
fn php_namespace_dynamic_and_global_resolution() {
    assert_str(
        "<?php\nnamespace A;\nfunction make(): string { return 'namespaced'; }\n\nnamespace {\n    function make(): string { return 'global'; }\n    echo \\A\\make() . '|' . make();\n}",
        "namespaced|global",
    );
}

#[test]
fn php_namespace_current_name_constant() {
    assert_str(
        "<?php\nnamespace Shop\\Payments;\necho __NAMESPACE__;\n",
        "Shop\\Payments",
    );
}

#[test]
fn php_namespace_aliasing_preferred_target() {
    assert_str(
        "<?php\nnamespace Core\\Util;\nclass Service {\n    public function id(): int { return 7; }\n}\n\nnamespace App;\nuse Core\\Util\\Service as CoreService;\necho (new CoreService())->id();\n",
        "7",
    );
}

#[test]
fn php_namespace_multiple_aliases_and_static_function_call() {
    assert_str(
        "<?php\nnamespace Src\\Lib;\nclass Fallback {\n    public static function name(): string { return 'class'; }\n}\nfunction action(): string { return 'fn'; }\n\nnamespace App;\nuse Src\\Lib\\Fallback as AliasClass;\nuse function Src\\Lib\\action as alias_fn;\n$kind = AliasClass::name();\necho $kind . '|' . alias_fn();\n",
        "class|fn",
    );
}

#[test]
fn php_namespace_nested_path_resolution() {
    assert_str(
        "<?php\nnamespace Company\\Module;\nclass Handler {\n    public function scope(): string { return 'module'; }\n}\n\nnamespace Company\\Module\\Sub;\nuse Company\\Module\\Handler;\necho (new Handler())->scope();\n",
        "module",
    );
}

#[test]
fn php_namespace_import_global_function_from_backslash() {
    assert_str(
        "<?php\nnamespace Demo;\nfunction trim_global(string $value): string { return 'local'; }\n\nnamespace {\n    use function Demo\\trim_global as local_trim;\n    echo local_trim('x') . '|' . \\trim(' x ');\n}\n",
        "local|x",
    );
}

#[test]
fn php_namespace_unqualified_fallback_to_current_namespace() {
    assert_str(
        "<?php\nnamespace Api\\Gateway;\nfunction route(): string { return 'inside'; }\nfunction run(): string { return route(); }\necho run();\n",
        "inside",
    );
}

#[test]
fn php_namespace_fully_qualified_class_resolution() {
    assert_str(
        "<?php\nnamespace Core;\nclass Logger { public function __construct() {} public function name(): string { return 'core'; } }\n\nnamespace App\\Http;\n$instance = new \\Core\\Logger();\necho $instance->name();\n",
        "core",
    );
}

#[test]
fn php_namespace_fully_qualified_function_resolution() {
    assert_str(
        "<?php\nnamespace Utils;\nfunction envTag(string $name): string { return \"util:$name\"; }\n\nnamespace App\\Runner;\necho \\Utils\\envTag('PAYMENTS');\n",
        "util:PAYMENTS",
    );
}

#[test]
fn php_namespace_imported_global_class_with_prefix_backslash() {
    assert_str(
        "<?php\nnamespace Tools;\nclass Box { public function tag(): string { return 'box'; } }\n\nnamespace {\n    use Tools\\Box;\n    $b = new \\Tools\\Box();\n    echo $b->tag();\n}\n",
        "box",
    );
}

#[test]
fn php_namespace_dynamic_name_runtime_works() {
    assert_str(
        "<?php\nnamespace Runtime;\nclass Service { public function endpoint(): string { return 'ok'; } }\n$name = 'Runtime\\\\Service';\n$class = new $name();\necho $class->endpoint();\n",
        "ok",
    );
}

#[test]
fn php_namespace_const_alias_in_unified_namespace() {
    assert_str(
        "<?php\nnamespace Billing;\nconst MODE = 'staging';\n\nnamespace App;\nuse const Billing\\MODE;\necho MODE . '|' . Billing\\MODE;\n",
        "staging|staging",
    );
}

#[test]
fn php_namespace_alias_conflicts_between_local_and_imported() {
    assert_str(
        "<?php\nnamespace Local;\nclass Worker {\n    public function env(): string { return 'local'; }\n}\n\nnamespace Runner;\nuse Local\\Worker;\nclass Worker {\n    public function env(): string { return 'imported'; }\n}\n$local = new \\Local\\Worker();\n$imported = new Worker();\necho $local->env() . '|' . $imported->env();\n",
        "local|imported",
    );
}

#[test]
fn php_namespace_group_use_style_runtime() {
    assert_str(
        "<?php\nnamespace Grouped\\Storage;\nclass FileStore { public static function kind(): string { return 'file'; } }\nfunction provider(): string { return 'provider'; }\nconst STORE = 'disk';\n\nnamespace App;\nuse Grouped\\Storage\\{FileStore, provider, STORE};\necho FileStore::kind() . '|' . provider() . '|' . STORE;\n",
        "file|provider|disk",
    );
}

#[test]
fn php_namespace_nested_use_function_alias_with_same_ns() {
    assert_str(
        "<?php\nnamespace Payments\\Stripe;\nfunction status(): string { return 'stripe'; }\n\nnamespace Payments;\nuse Stripe\\status as stripe_status;\necho stripe_status();\n",
        "stripe",
    );
}

#[test]
fn php_namespace_relative_resolution_after_import() {
    assert_str(
        "<?php\nnamespace Infra\\Storage;\nclass Disk { public function label(): string { return 'disk'; } }\n\nnamespace App;\nuse Infra\\Storage\\Disk;\necho (new Disk())->label();\n",
        "disk",
    );
}

#[test]
fn php_namespace_function_exists_for_qualified_names() {
    assert_str(
        "<?php\nnamespace Util\\Io;\nfunction open_file(): bool { return true; }\n\nnamespace App;\necho (function_exists('Util\\\\Io\\\\open_file') ? 'yes' : 'no');\necho '|';\necho (\\function_exists('Util\\\\Io\\\\open_file') ? 'yes' : 'no');\n",
        "yes|yes",
    );
}

#[test]
fn php_namespace_trait_import_in_namespace() {
    assert_str(
        "<?php\nnamespace Core\\Helpers;\ntrait Formatter {\n    public function format(string $v): string { return 'fmt:' . $v; }\n}\n\nnamespace App;\nuse Core\\Helpers\\Formatter;\nclass Report {\n    use Formatter;\n}\necho (new Report())->format('x');\n",
        "fmt:x",
    );
}

#[test]
fn php_namespace_backed_enum_in_namespace_runtime() {
    assert_str(
        "<?php\nnamespace App\\Domain;\nenum Mode: string {\n    case Read = 'r';\n    case Write = 'w';\n}\n\necho Mode::Read->value;\n",
        "r",
    );
}

#[test]
fn php_namespace_dynamic_namespace_string() {
    assert_str(
        "<?php\nnamespace RuntimeNs;\nclass Service { public function run(): string { return 'run'; } }\n$ns = __NAMESPACE__ . '\\\\';\n$name = $ns . 'Service';\n$instance = new $name();\necho $instance->run();\n",
        "run",
    );
}

#[test]
fn php_namespace_class_alias_preserves_static_property() {
    assert_str(
        "<?php\nnamespace BaseNs;\nclass Counter { public static int $n = 1; public static function get(): int { return self::$n; } }\n\nnamespace App;\nuse BaseNs\\Counter as C;\necho C::$n;\necho '|';\necho C::get();\n",
        "1|1",
    );
}

#[test]
fn php_namespace_relative_resolution_with_class_exists() {
    assert_str(
        r#"<?php
namespace Framework\Core;
class Kernel {
    public static function name(): string { return 'kernel'; }
}

namespace App\Runtime;
echo class_exists('Kernel') ? 'inner' : 'miss';
echo '|';
echo class_exists('\\Framework\\Core\\Kernel') ? 'absolute' : 'noabs';
echo '|';
echo \Framework\Core\Kernel::name();
"#,
        "miss|absolute|kernel",
    );
}

#[test]
fn php_namespace_dynamic_function_call_across_namespaces() {
    assert_str(
        r#"<?php
namespace Tools {
    function format(string $s): string { return 'fmt:' . $s; }
}

namespace App {
    $call = '\\Tools\\format';
    echo call_user_func($call, 'x');
}"#,
        "fmt:x",
    );
}

#[test]
fn php_namespace_function_exists_for_current_namespace() {
    assert_str(
        r#"<?php
namespace Local {
    function active(): bool { return true; }
}

namespace {
    use function Local\active;
    echo function_exists('active') ? 'no' : 'yes';
    echo '|';
    echo is_callable('Local\\active') ? 'callable' : 'nocal';
}
"#,
        "yes|callable",
    );
}

#[test]
fn php_namespace_backed_class_alias_runtime() {
    assert_str(
        r#"<?php
namespace Infra\Models;
class Entity {
    public function id(): int { return 42; }
}
class_alias(Entity::class, 'GlobalEntity');

namespace App;
echo class_exists('GlobalEntity') ? (new GlobalEntity())->id() : 0;
"#,
        "42",
    );
}

#[test]
fn php_namespace_trait_aliased_import_runtime() {
    assert_str(
        r#"<?php
namespace Traits;
trait Logger {
    public function log(): string { return 'ok'; }
}

namespace App;
use Traits\Logger as AppLogger;
class Service {
    use AppLogger { log as public emit; }
}
echo (new Service())->emit();
"#,
        "ok",
    );
}
