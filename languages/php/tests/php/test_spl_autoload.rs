//! `spl_autoload_register` and class/interface/trait existence checks.

crate::php_cases! {
    spl_autoload_register_loads_namespaced_class => {
        r#"<?php
spl_autoload_register(function (string $class): void {
    if ($class === 'App\\Widget') {
        eval('namespace App; class Widget { public function id(): string { return "w1"; } }');
    }
});
echo (new App\Widget())->id();
"#,
        ["w1"]
    };

    spl_autoload_unregister_leaves_class_loaded => {
        r#"<?php
$loader = function (string $class): void {
    if ($class === 'Tmp\\Once') {
        eval('namespace Tmp; class Once {}');
    }
};
spl_autoload_register($loader);
class_exists('Tmp\\Once');
spl_autoload_unregister($loader);
echo class_exists('Tmp\\Once', false) ? 'loaded' : 'gone';
"#,
        ["loaded"]
    };

    class_exists_without_autoload_returns_false_for_missing => {
        r#"<?php
echo class_exists('Missing\\Class', false) ? 'yes' : 'no';
"#,
        ["no"]
    };

    class_exists_after_definition_returns_true => {
        r#"<?php
class LocalSvc {}
echo class_exists('LocalSvc', false) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    interface_exists_detects_declared_interface => {
        r#"<?php
interface Capable { public function run(): void; }
echo interface_exists('Capable', false) ? 'iface' : 'no';
"#,
        ["iface"]
    };

    trait_exists_detects_declared_trait => {
        r#"<?php
trait Loggable { public function log(): string { return 'log'; } }
echo trait_exists('Loggable', false) ? 'trait' : 'no';
"#,
        ["trait"]
    };

    method_exists_on_instance => {
        r#"<?php
class Svc { public function handle(): void {} }
echo method_exists(new Svc(), 'handle') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    method_exists_missing_returns_false => {
        r#"<?php
class Svc {}
echo method_exists('Svc', 'missing') ? 'yes' : 'no';
"#,
        ["no"]
    };

    property_exists_checks_declared_property => {
        r#"<?php
class Box { public string $label = 'x'; }
echo property_exists(new Box(), 'label') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    get_declared_classes_includes_stdclass => {
        r#"<?php
echo in_array('stdClass', get_declared_classes(), true) ? 'std' : 'no';
"#,
        ["std"]
    };

    get_declared_interfaces_includes_stringable => {
        r#"<?php
echo interface_exists('Stringable', false) ? 'str' : 'no';
"#,
        ["str"]
    };

    get_declared_traits_lists_user_trait => {
        r#"<?php
trait Marker {}
echo in_array('Marker', get_declared_traits(), true) ? 'marked' : 'no';
"#,
        ["marked"]
    };

    autoload_stack_invokes_registered_loader => {
        r#"<?php
$hit = false;
spl_autoload_register(function (string $class) use (&$hit): void {
    if ($class === 'Hit\\Load') { $hit = true; }
});
class_exists('Hit\\Load');
echo $hit ? 'hit' : 'miss';
"#,
        ["hit"]
    };

    class_alias_creates_alias_name => {
        r#"<?php
class RealThing {}
class_alias(RealThing::class, 'AliasThing');
echo (new AliasThing()) instanceof RealThing ? 'alias' : 'no';
"#,
        ["alias"]
    };

    get_parent_class_reports_extends => {
        r#"<?php
class ParentCls {}
class ChildCls extends ParentCls {}
echo get_parent_class(ChildCls::class);
"#,
        ["ParentCls"]
    };

    is_subclass_of_detects_inheritance => {
        r#"<?php
class Base {}
class Derived extends Base {}
echo is_subclass_of(Derived::class, Base::class) ? 'sub' : 'no';
"#,
        ["sub"]
    };

    is_a_with_string_class_name => {
        r#"<?php
class Node {}
$n = new Node();
echo is_a($n, 'Node') ? 'node' : 'other';
"#,
        ["node"]
    };

    spl_autoload_register_accepts_static_call => {
        r#"<?php
class Loader {
    public static function init(string $class): void {
        if ($class === 'Auto\\Tool') {
            eval('namespace Auto; class Tool { public function label(): string { return "tool"; } }');
        }
    }
}
spl_autoload_register([Loader::class, 'init']);
echo (new Auto\Tool())->label();
"#,
        ["tool"]
    };

    autoload_ignores_redundant_registrations => {
        r#"<?php
spl_autoload_register(function(string $class): void {
    if ($class === 'Dup\\One') {
        eval('namespace Dup; class One {}');
    }
});
spl_autoload_register(function(string $class): void {
    if ($class === 'Dup\\One') {
        eval('namespace Dup; class One {}');
    }
});
echo class_exists('Dup\\One') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    autoload_order_preserves_first_match => {
        r#"<?php
$hits = [];
spl_autoload_register(function(string $class) use (&$hits): void {
    if ($class === 'Chain\\Svc') { $hits[] = 'first'; eval('namespace Chain; class Svc {}'); }
});
spl_autoload_register(function(string $class) use (&$hits): void {
    if ($class === 'Chain\\Svc') { $hits[] = 'second'; }
});
class_exists('Chain\\Svc');
echo implode('|', $hits);
"#,
        ["first"]
    };

    autoload_unregister_removes_loader => {
        r#"<?php
$count = 0;
$loader = function(string $class) use (&$count): void {
    if ($class === 'Temp\\Svc') {
        $count++;
        eval('namespace Temp; class Svc {}');
    }
};
spl_autoload_register($loader);
spl_autoload_unregister($loader);
class_exists('Temp\\Svc');
echo $count;
"#,
        ["0"]
    };

    class_parents_with_parent => {
        r#"<?php
class P {}
class C extends P {}
$parents = class_parents(C::class);
echo in_array('P', $parents, true) ? 'has' : 'no';
"#,
        ["has"]
    };

    class_implements_with_interface => {
        r#"<?php
interface Contract {}
class Impl implements Contract {}
$interfaces = class_implements(Impl::class);
echo in_array('Contract', $interfaces, true) ? 'ok' : 'no';
"#,
        ["ok"]
    };

    get_object_vars_lists_public_fields => {
        r#"<?php
class Holder {
    public $a = 1;
    private $b = 2;
    protected $c = 3;
}
$h = new Holder();
$vars = get_object_vars($h);
echo array_key_exists('a', $vars) ? 'a' : 'na';
echo array_key_exists('b', $vars) ? '|b' : '|nb';
echo array_key_exists('c', $vars) ? '|c' : '|nc';
"#,
        ["a|nb|nc"]
    };

    class_alias_name_collision_returns_false => {
        r#"<?php
class BaseThing {}
class_alias(BaseThing::class, 'AliasThing2');
$ok = class_alias(BaseThing::class, 'AliasThing2', false);
echo $ok ? 'created' : 'failed';
"#,
        ["failed"]
    };

    interface_with_traits_list_via_check => {
        r#"<?php
interface I {}
echo interface_exists('I', false) ? 'yes' : 'no';
echo '|';
echo trait_exists('I', false) ? 'yes' : 'no';
"#,
        ["yes|no"]
    };

    spl_autoload_register_invokable_loader => {
        r#"<?php
class InvokableLoader {
    public function __invoke(string $class): void {
        if ($class === 'Invokable\\Service') {
            eval('namespace Invokable; class Service { public function id(): string { return \"svc\"; } }');
        }
    }
}
spl_autoload_register(new InvokableLoader());
$svc = new Invokable\Service();
echo $svc->id();
"#,
        ["svc"]
    };

    spl_autoload_register_prepend_controls_order => {
        r#"<?php
$log = [];
spl_autoload_register(function (string $class) use (&$log): void {
    if ($class === 'Order\\Widget') {
        $log[] = 'base';
        eval('namespace Order; class Widget {}');
    }
});
spl_autoload_register(function (string $class) use (&$log): void {
    if ($class === 'Order\\Widget') {
        $log[] = 'prepended';
    }
}, true, true);

if (class_exists('Order\\Widget')) {
    echo implode(',', $log) . '|loaded';
} else {
    echo implode(',', $log) . '|missing';
}
"#,
        ["prepended,base|loaded"]
    };

    spl_autoload_call_invokes_registered_loaders => {
        r#"<?php
$log = [];
$loader = function (string $class) use (&$log): void {
    if ($class === 'Manual\\Probe') {
        $log[] = 'called';
        eval('namespace Manual; class Probe {}');
    }
};
spl_autoload_register($loader);
spl_autoload_call('Manual\\Probe');
echo (class_exists('Manual\\Probe', false) ? 'exists' : 'missing') . '|' . implode(',', $log);
"#,
        ["exists|called"]
    };

    spl_autoload_functions_tracks_added_handler => {
        r#"<?php
$before = function_exists('spl_autoload_functions') ? count((array) spl_autoload_functions()) : 0;
$loader = function (string $class): void {
    if ($class === 'Unused\\Class') {
        eval('namespace Unused; class Class {}');
    }
};
spl_autoload_register($loader);
$afterRegister = count((array) spl_autoload_functions());
spl_autoload_unregister($loader);
$afterUnregister = count((array) spl_autoload_functions());
echo ($afterRegister === $before + 1 ? 'plus1' : 'no');
echo '|';
echo ($afterUnregister === $before ? 'clean' : 'dirty');
"#,
        ["plus1|clean"]
    };

    interface_exists_can_autoload_from_callback => {
        r#"<?php
$loader = function (string $class): void {
    if ($class === 'Auto\\IInspectable') {
        eval('namespace Auto; interface IInspectable { public function ok(): bool; }');
    }
};
spl_autoload_register($loader);
echo interface_exists('Auto\\IInspectable') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    trait_exists_can_autoload_from_callback => {
        r#"<?php
$loader = function (string $class): void {
    if ($class === 'Auto\\TInspectable') {
        eval('namespace Auto; trait TInspectable {}');
    }
};
spl_autoload_register($loader);
echo trait_exists('Auto\\TInspectable') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    method_exists_reports_case_insensitive => {
        r#"<?php
class Worker {
    public function handlePayload(): string { return 'ok'; }
}
$obj = new Worker();
echo method_exists($obj, 'handlepayload') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    property_exists_on_class_name_sees_private => {
        r#"<?php
class Holder {
    private string $secret = 'x';
}
echo property_exists(Holder::class, 'secret') ? 'yes' : 'no';
echo '|';
echo property_exists(new Holder(), 'secret') ? 'yes' : 'no';
"#,
        ["yes|yes"]
    };

    class_exists_false_skips_autoload_callback => {
        r#"<?php
$hit = 0;
$loader = function (string $class) use (&$hit): void {
    $hit++;
    if ($class === 'Skip\\Target') {
        eval('namespace Skip; class Target {}');
    }
};
spl_autoload_register($loader);
$result = class_exists('Skip\\Target', false) ? 'hit' : 'skip';
echo $result . '|' . $hit;
spl_autoload_unregister($loader);
"#,
        ["skip|0"]
    };

    spl_autoload_reports_loader_argument_exact => {
        r#"<?php
$classes = [];
spl_autoload_register(function (string $class): void use (&$classes) {
    $classes[] = $class;
    if ($class === 'Exact\\Class') {
        eval('namespace Exact; class Class {}');
    }
});
class_exists('Exact\\Class');
echo $classes[0] === 'Exact\\Class' ? 'exact' : 'miss';
"#,
        ["exact"]
    };

    spl_autoload_prepend_false_keeps_tail_position => {
        r#"<?php
$hit = [];
spl_autoload_register(function (string $class) use (&$hit): void {
    if ($class === 'Pre\\Service') { $hit[] = 'first'; eval('namespace Pre; class Service {}'); }
});
$loader = function (string $class) use (&$hit): void {
    if ($class === 'Pre\\Service') { $hit[] = 'second'; }
};
spl_autoload_register($loader, true, false);
class_exists('Pre\\Service');
echo implode('|', $hit);
"#,
        ["first|second"]
    };

    spl_autoload_prepend_true_puts_loader_before_existing => {
        r#"<?php
$hit = [];
spl_autoload_register(function (string $class) use (&$hit): void {
    if ($class === 'Prepend\\Thing') { $hit[] = 'tail'; eval('namespace Prepend; class Thing {}'); }
});
$loader = function (string $class) use (&$hit): void {
    if ($class === 'Prepend\\Thing') { $hit[] = 'head'; }
};
spl_autoload_register($loader, true, true);
class_exists('Prepend\\Thing');
echo implode('|', $hit);
"#,
        ["head|tail"]
    };

    class_alias_reuses_existing_when_preferred => {
        r#"<?php
class Proto {}
class_alias(Proto::class, 'AliasProto');
echo (new AliasProto()) instanceof Proto ? 'yes' : 'no';
echo '|';
echo is_subclass_of(AliasProto::class, Proto::class) ? 'sub' : 'no';
"#,
        ["yes|no"]
    };

    interface_implements_chain_includes_all => {
        r#"<?php
interface Root {}
interface Mid extends Root {}
interface Leaf extends Mid {}
class Impl implements Leaf {}
$interfaces = class_implements(Impl::class);
echo in_array('Root', $interfaces, true) ? 'root' : 'n1';
echo '|';
echo in_array('Mid', $interfaces, true) ? 'mid' : 'n2';
echo '|';
echo in_array('Leaf', $interfaces, true) ? 'leaf' : 'n3';
"#,
        ["root|mid|leaf"]
    };

    autoload_call_count_when_class_loaded_once => {
        r#"<?php
$calls = 0;
spl_autoload_register(function (string $class) use (&$calls): void {
    if ($class === 'Cache\\Svc') {
        $calls++;
        eval('namespace Cache; class Svc { public function name(): string { return \"svc\"; } }');
    }
});
class_exists('Cache\\Svc');
class_exists('Cache\\Svc', false);
class_exists('Cache\\Svc');
echo $calls;
"#,
        ["1"]
    };

interface_autoload_triggered_on_class_of_existing_type => {
        r#"<?php
$calls = [];
spl_autoload_register(function (string $class) use (&$calls): void {
    if ($class === 'AutoLoad\\IWidget') {
        $calls[] = $class;
        eval('namespace AutoLoad; interface IWidget { public function render(): void; }');
    }
});
interface_exists('AutoLoad\\IWidget');
echo count($calls);
"#,
        ["1"]
    };

    trait_autoload_only_once_per_name => {
        r#"<?php
$seen = 0;
spl_autoload_register(function (string $class) use (&$seen): void {
    if ($class === 'AutoLoad\\TWidget') {
        $seen += 1;
        if ($seen === 1) {
            eval('namespace AutoLoad; trait TWidget {}');
        }
    }
});
trait_exists('AutoLoad\\TWidget');
trait_exists('AutoLoad\\TWidget');
echo $seen;
"#,
        ["1"]
    };

    invokable_object_loader_receives_class_name => {
        r#"<?php
class Collector {
    public array $seen = [];
    public function __invoke(string $class): void {
        $this->seen[] = $class;
        if ($class === 'AutoLoad\\Loaded') {
            eval('namespace AutoLoad; class Loaded {}');
        }
    }
}
$loader = new Collector();
spl_autoload_register($loader);
class_exists('AutoLoad\\Loaded');
echo $loader->seen[0] === 'AutoLoad\\Loaded' ? 'ok' : 'bad';
"#,
        ["ok"]
    };

    get_declared_classes_reflects_autoloaded => {
        r#"<?php
spl_autoload_register(function (string $class): void {
    if ($class === 'AutoLoad\\Transient') {
        eval('namespace AutoLoad; class Transient {}');
    }
});
class_exists('AutoLoad\\Transient');
$declared = get_declared_classes();
echo in_array('AutoLoad\\Transient', $declared, true) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    autoload_function_name_lookup_with_string_callable => {
        r#"<?php
function spl_autoload_string_loader(string $class): void {
    if ($class === 'Fn\\Service') {
        eval('namespace Fn; class Service { public function value(): string { return \"fn\"; } }');
    }
}
spl_autoload_register('spl_autoload_string_loader');
echo (new Fn\Service())->value();
"#,
        ["fn"]
    };

    spl_autoload_register_with_array_callback => {
        r#"<?php
class ArrayLoader {
    public static function load(string $class): void {
        if ($class === 'Array\\Svc') {
            eval('namespace Array; class Svc { public function name(): string { return \"array\"; } }');
        }
    }
}
spl_autoload_register([ArrayLoader::class, 'load']);
echo (new Array\Svc())->name();
"#,
        ["array"]
    };

    class_uses_includes_trait_after_use => {
        r#"<?php
trait UsesA {}
class UsesContainer { use UsesA; }
echo in_array('UsesA', class_uses(UsesContainer::class), true) ? 'used' : 'no';
"#,
        ["used"]
    };

    class_uses_runtime_chain => {
        r#"<?php
trait T1 {}
trait T2 { use T1; }
class C { use T2; }
echo count(class_uses(C::class, true)) === 2 ? 'chain' : 'no';
"#,
        ["chain"]
    };

    trait_uses_recursive_nested => {
        r#"<?php
trait UnitA {}
trait UnitB { use UnitA; }
echo count(trait_uses(UnitB::class, true)) === 1 ? 'yes' : 'no';
"#,
        ["yes"]
    };

    class_implements_multiple_interfaces => {
        r#"<?php
interface IA {}
interface IB {}
class Multi implements IA, IB {}
$interfaces = class_implements(Multi::class, true);
echo in_array('IA', $interfaces, true) ? 'ia' : 'no';
echo '|';
echo in_array('IB', $interfaces, true) ? 'ib' : 'no';
"#,
        ["ia|ib"]
    };

    interface_exists_case_sensitive_false => {
        r#"<?php
interface NameCase {}
echo interface_exists('namecase') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    trait_exists_case_sensitive_false => {
        r#"<?php
trait NameTrait {}
echo trait_exists('nametrait') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_a_accepts_instance_object => {
        r#"<?php
class BaseClass {}
$obj = new BaseClass();
echo is_a($obj, 'BaseClass') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    get_parent_class_with_instance_argument => {
        r#"<?php
class A {}
class B extends A {}
echo get_parent_class(new B()) === 'A' ? 'parent' : 'no';
"#,
        ["parent"]
    };

    class_alias_can_access_static_methods => {
        r#"<?php
class OriginalService {
    public static function marker(): string { return 'ok'; }
}
class_alias(OriginalService::class, 'AliasService');
echo AliasService::marker();
"#,
        ["ok"]
    };

    class_alias_keeps_interfaces_and_inheritance => {
        r#"<?php
interface IBase {}
class Source implements IBase {}
class_alias(Source::class, 'AliasSource');
echo is_subclass_of('AliasSource', 'IBase') ? 'implements' : 'no';
"#,
        ["implements"]
    };

    spl_autoload_call_for_missing_class_returns_null => {
        r#"<?php
spl_autoload_register(function (string $class): void {
    if ($class === 'Noop\\Known') {
        eval('namespace Noop; class Known {}');
    }
});
$result = spl_autoload_call('Missing\\NeverFound');
echo is_null($result) ? 'null' : 'value';
echo '|';
echo class_exists('Noop\\Known') ? 'known' : 'no';
"#,
        ["null|known"]
    };
}
