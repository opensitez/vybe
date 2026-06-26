//! noSuchMethod: proxy forwarding, Invocation metadata, dynamic dispatch,
//! and default Object.toString when methods are absent.

dart_cases! {
    no_such_method_returns_constant_for_missing_getter => {
        r#"class Proxy {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 99;
  }
}
void main() {
  dynamic p = Proxy();
  print(p.missing);
}"#,
        ["99"]
    };

    no_such_method_returns_constant_for_missing_method => {
        r#"class Proxy {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 42;
  }
}
void main() {
  dynamic p = Proxy();
  print(p.missing());
}"#,
        ["42"]
    };

    no_such_method_prints_member_name_symbol => {
        r#"class Logger {
  @override
  dynamic noSuchMethod(Invocation inv) {
    print(inv.memberName.toString());
    return 0;
  }
}
void main() {
  dynamic l = Logger();
  l.fetchData();
}"#,
        ["Symbol(\"fetchData\")", "0"]
    };

    no_such_method_detects_method_invocation => {
        r#"class Probe {
  @override
  dynamic noSuchMethod(Invocation inv) {
    print(inv.isMethod);
    return 0;
  }
}
void main() {
  dynamic p = Probe();
  p.run();
}"#,
        ["true", "0"]
    };

    no_such_method_detects_getter_invocation => {
        r#"class Probe {
  @override
  dynamic noSuchMethod(Invocation inv) {
    print(inv.isGetter);
    return 0;
  }
}
void main() {
  dynamic p = Probe();
  print(p.value);
}"#,
        ["true", "0"]
    };

    no_such_method_detects_setter_invocation => {
        r#"class Probe {
  @override
  dynamic noSuchMethod(Invocation inv) {
    print(inv.isSetter);
    return 0;
  }
}
void main() {
  dynamic p = Probe();
  p.value = 1;
  print(0);
}"#,
        ["true", "0"]
    };

    no_such_method_reads_positional_arguments => {
        r#"class Echo {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.positionalArguments.length;
  }
}
void main() {
  dynamic e = Echo();
  print(e.sum(1, 2, 3));
}"#,
        ["3"]
    };

    no_such_method_sums_positional_arguments => {
        r#"class Adder {
  @override
  dynamic noSuchMethod(Invocation inv) {
    var total = 0;
    for (var a in inv.positionalArguments) {
      total = total + (a as int);
    }
    return total;
  }
}
void main() {
  dynamic a = Adder();
  print(a.add(10, 20, 30));
}"#,
        ["60"]
    };

    no_such_method_reads_named_arguments_map => {
        r#"class Named {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.namedArguments.length;
  }
}
void main() {
  dynamic n = Named();
  print(n.config(x: 1, y: 2));
}"#,
        ["2"]
    };

    no_such_method_extracts_named_argument_value => {
        r#"class Named {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.namedArguments[#mode];
  }
}
void main() {
  dynamic n = Named();
  print(n.run(mode: 'fast'));
}"#,
        ["fast"]
    };

    no_such_method_forwards_to_target_map => {
        r#"class Forwarder {
  final Map<String, dynamic> target;
  Forwarder(this.target);
  @override
  dynamic noSuchMethod(Invocation inv) {
    var name = inv.memberName.toString();
    if (name.contains('get')) {
      var key = name.replaceAll('Symbol(\"', '').replaceAll('\")', '');
      return target[key];
    }
    return null;
  }
}
void main() {
  dynamic f = Forwarder({'x': 7});
  print(f.x);
}"#,
        ["7"]
    };

    no_such_method_proxy_delegates_method_to_target => {
        r#"class Target {
  int doubleIt(int n) {
    return n * 2;
  }
}
class Proxy {
  final Target target;
  Proxy(this.target);
  @override
  dynamic noSuchMethod(Invocation inv) {
    if (inv.isMethod && inv.memberName == #doubleIt) {
      return target.doubleIt(inv.positionalArguments[0] as int);
    }
    return super.noSuchMethod(inv);
  }
}
void main() {
  dynamic p = Proxy(Target());
  print(p.doubleIt(5));
}"#,
        ["10"]
    };

    no_such_method_dynamic_call_chained => {
        r#"class Chain {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return Chain();
  }
  int end() {
    return 1;
  }
}
void main() {
  dynamic c = Chain();
  print(c.next().end());
}"#,
        ["1"]
    };

    no_such_method_returns_string_from_method_name => {
        r#"class NameEcho {
  @override
  dynamic noSuchMethod(Invocation inv) {
    var s = inv.memberName.toString();
    return s.replaceAll('Symbol(\"', '').replaceAll('\")', '');
  }
}
void main() {
  dynamic n = NameEcho();
  print(n.hello());
}"#,
        ["hello"]
    };

    no_such_method_zero_arg_method => {
        r#"class Zero {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.positionalArguments.isEmpty;
  }
}
void main() {
  dynamic z = Zero();
  print(z.ping());
}"#,
        ["true"]
    };

    no_such_method_first_positional_argument => {
        r#"class First {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.positionalArguments.first;
  }
}
void main() {
  dynamic f = First();
  print(f.take('alpha'));
}"#,
        ["alpha"]
    };

    no_such_method_last_positional_argument => {
        r#"class Last {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.positionalArguments.last;
  }
}
void main() {
  dynamic l = Last();
  print(l.pick(1, 2, 9));
}"#,
        ["9"]
    };

    default_to_string_on_plain_class => {
        r#"class Plain {}
void main() {
  print(Plain().toString().contains('Plain'));
}"#,
        ["true"]
    };

    default_to_string_includes_instance_prefix => {
        r#"class Widget {}
void main() {
  print(Widget().toString().startsWith('Instance of'));
}"#,
        ["true"]
    };

    no_such_method_custom_to_string_still_works => {
        r#"class Labeled {
  @override
  String toString() {
    return 'labeled';
  }
}
void main() {
  print(Labeled().toString());
}"#,
        ["labeled"]
    };

    no_such_method_does_not_intercept_declared_method => {
        r#"class Mixed {
  int real() {
    return 1;
  }
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 0;
  }
}
void main() {
  print(Mixed().real());
}"#,
        ["1"]
    };

    no_such_method_does_not_intercept_declared_getter => {
        r#"class Mixed {
  int get value {
    return 5;
  }
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 0;
  }
}
void main() {
  print(Mixed().value);
}"#,
        ["5"]
    };

    dynamic_dispatch_through_no_such_method => {
        r#"class Dyn {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 'dyn';
  }
}
void main() {
  dynamic d = Dyn();
  print(d.anything());
}"#,
        ["dyn"]
    };

    dynamic_variable_invokes_missing_method => {
        r#"class Handler {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.positionalArguments[0];
  }
}
void main() {
  var obj = Handler();
  dynamic d = obj;
  print(d.wrap(88));
}"#,
        ["88"]
    };

    no_such_method_proxy_records_call_count => {
        r#"class Counter {
  int calls = 0;
  @override
  dynamic noSuchMethod(Invocation inv) {
    calls++;
    return calls;
  }
}
void main() {
  dynamic c = Counter();
  c.a();
  c.b();
  print(c.calls);
}"#,
        ["2"]
    };

    no_such_method_returns_bool_for_predicate_name => {
        r#"class Pred {
  @override
  dynamic noSuchMethod(Invocation inv) {
    var name = inv.memberName.toString();
    return name.contains('is');
  }
}
void main() {
  dynamic p = Pred();
  print(p.isReady());
}"#,
        ["true"]
    };

    no_such_method_dynamic_method_returns_list_element => {
        r#"class ListProxy {
  final List<int> data = [10, 20, 30];
  @override
  dynamic noSuchMethod(Invocation inv) {
    if (inv.isMethod && inv.memberName == #at) {
      return data[inv.positionalArguments[0] as int];
    }
    return null;
  }
}
void main() {
  dynamic p = ListProxy();
  print(p.at(1));
}"#,
        ["20"]
    };

    no_such_method_super_fallback_not_used_when_handled => {
        r#"class Safe {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 'handled';
  }
}
void main() {
  dynamic s = Safe();
  print(s.missing());
}"#,
        ["handled"]
    };

    no_such_method_with_multiple_named_args => {
        r#"class Config {
  @override
  dynamic noSuchMethod(Invocation inv) {
    var a = inv.namedArguments[#a] as int;
    var b = inv.namedArguments[#b] as int;
    return a + b;
  }
}
void main() {
  dynamic c = Config();
  print(c.merge(a: 3, b: 4));
}"#,
        ["7"]
    };

    no_such_method_empty_named_arguments => {
        r#"class EmptyNamed {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.namedArguments.isEmpty;
  }
}
void main() {
  dynamic e = EmptyNamed();
  print(e.go(1));
}"#,
        ["true"]
    };

    no_such_method_returns_list_from_args => {
        r#"class Collect {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.positionalArguments.join('-');
  }
}
void main() {
  dynamic c = Collect();
  print(c.join('a', 'b', 'c'));
}"#,
        ["a-b-c"]
    };

    no_such_method_on_subclass_inherits_override => {
        r#"class Base {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 1;
  }
}
class Sub extends Base {}
void main() {
  dynamic s = Sub();
  print(s.m());
}"#,
        ["1"]
    };

    no_such_method_subclass_overrides_parent => {
        r#"class Base {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 1;
  }
}
class Sub extends Base {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 2;
  }
}
void main() {
  dynamic s = Sub();
  print(s.m());
}"#,
        ["2"]
    };

    no_such_method_dynamic_setter_side_effect => {
        r#"class Store {
  int stored = 0;
  @override
  dynamic noSuchMethod(Invocation inv) {
    if (inv.isSetter) {
      stored = inv.positionalArguments[0] as int;
    }
    return null;
  }
}
void main() {
  dynamic s = Store();
  s.value = 15;
  print(s.stored);
}"#,
        ["15"]
    };

    no_such_method_member_name_for_getter => {
        r#"class Tag {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.memberName == #title;
  }
}
void main() {
  dynamic t = Tag();
  print(t.title);
}"#,
        ["true"]
    };

    no_such_method_member_name_for_method => {
        r#"class Tag {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.memberName == #run;
  }
}
void main() {
  dynamic t = Tag();
  print(t.run());
}"#,
        ["true"]
    };

    no_such_method_forwarding_wrapper_pattern => {
        r#"class Real {
  int add(int a, int b) {
    return a + b;
  }
}
class Wrapper {
  final Real inner;
  Wrapper(this.inner);
  @override
  dynamic noSuchMethod(Invocation inv) {
    if (inv.memberName == #add) {
      return inner.add(
        inv.positionalArguments[0] as int,
        inv.positionalArguments[1] as int,
      );
    }
    return null;
  }
}
void main() {
  dynamic w = Wrapper(Real());
  print(w.add(2, 3));
}"#,
        ["5"]
    };

    no_such_method_returns_type_string => {
        r#"class TypeEcho {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.isMethod ? 'method' : 'other';
  }
}
void main() {
  dynamic t = TypeEcho();
  print(t.f());
}"#,
        ["method"]
    };

    dynamic_call_after_assignment => {
        r#"class Box {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 100;
  }
}
void main() {
  dynamic b;
  b = Box();
  print(b.fetch());
}"#,
        ["100"]
    };

    no_such_method_with_string_return_from_dynamic => {
        r#"class StrProxy {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 'proxy-${inv.positionalArguments.length}';
  }
}
void main() {
  dynamic s = StrProxy();
  print(s.msg('a', 'b'));
}"#,
        ["proxy-2"]
    };

    no_such_method_isMethod_false_for_getter => {
        r#"class Check {
  @override
  dynamic noSuchMethod(Invocation inv) {
    print(inv.isMethod);
    return 0;
  }
}
void main() {
  dynamic c = Check();
  print(c.field);
}"#,
        ["false", "0"]
    };

    no_such_method_isGetter_false_for_method => {
        r#"class Check {
  @override
  dynamic noSuchMethod(Invocation inv) {
    print(inv.isGetter);
    return 0;
  }
}
void main() {
  dynamic c = Check();
  c.run();
}"#,
        ["false", "0"]
    };

    default_hashcode_on_object => {
        r#"class Empty {}
void main() {
  var a = Empty();
  var b = Empty();
  print(a.hashCode == b.hashCode);
}"#,
        ["false"]
    };

    no_such_method_preserves_declared_to_string => {
        r#"class Named {
  @override
  String toString() {
    return 'named';
  }
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 0;
  }
}
void main() {
  dynamic n = Named();
  print(n.toString());
}"#,
        ["named"]
    };

    no_such_method_double_dynamic_dispatch => {
        r#"class A {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return B();
  }
}
class B {
  int val() {
    return 3;
  }
}
void main() {
  dynamic a = A();
  print(a.next().val());
}"#,
        ["3"]
    };

    no_such_method_positional_args_types => {
        r#"class Types {
  @override
  dynamic noSuchMethod(Invocation inv) {
    var a = inv.positionalArguments[0];
    var b = inv.positionalArguments[1];
    return '$a:$b';
  }
}
void main() {
  dynamic t = Types();
  print(t.pair(1, 'x'));
}"#,
        ["1:x"]
    };

    no_such_method_named_only_call => {
        r#"class OnlyNamed {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.namedArguments[#key];
  }
}
void main() {
  dynamic o = OnlyNamed();
  print(o.lookup(key: 'dart'));
}"#,
        ["dart"]
    };
}
