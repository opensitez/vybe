use crate::helpers::run_prints;

#[test]
fn test_interface_declaration() {
    let out = run_prints(
        r#"
        interface Printable {
            fun printMsg()
        }

        class MessagePrinter : Printable {
            fun printMsg() {
                println("interface message")
            }
        }

        fun main() {
            val printer = MessagePrinter()
            printer.printMsg()
        }
    "#,
    );
    assert_eq!(out, &["interface message"]);
}

#[test]
fn test_multiple_interface_implementation() {
    let out = run_prints(
        r#"
        interface Named {
            fun getName(): String
        }

        interface Aged {
            fun getAge(): Int
        }

        class Citizen(val name: String, val age: Int) : Named, Aged {
            override fun getName(): String = name
            override fun getAge(): Int = age
        }

        fun main() {
            val c = Citizen("Bob", 40)
            println(c.getName())
            println(c.getAge())
        }
    "#,
    );
    assert_eq!(out, &["Bob", "40"]);
}

#[test]
fn test_interface_default_method() {
    let out = run_prints(
        r#"
        interface Notifier {
            fun notify(): String {
                return "notified"
            }
        }

        class SilentNotifier : Notifier

        fun main() {
            val n: Notifier = SilentNotifier()
            println(n.notify())
        }
    "#,
    );
    assert_eq!(out, &["notified"]);
}

#[test]
fn test_interface_property_contract() {
    let out = run_prints(
        r#"
        interface Identified {
            val id: Int
        }

        class Item(override val id: Int) : Identified

        fun main() {
            val item: Identified = Item(7)
            println(item.id)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_interface_inheritance() {
    let out = run_prints(
        r#"
        interface Named {
            fun name(): String
        }

        interface Described : Named {
            fun description(): String
        }

        class Product : Described {
            override fun name(): String = "p"
            override fun description(): String = name() + "rod"
        }

        fun main() {
            val p = Product()
            println(p.name())
            println(p.description())
        }
    "#,
    );
    assert_eq!(out, &["p", "prod"]);
}

#[test]
fn test_interface_default_and_override_method() {
    let out = run_prints(
        r#"
        interface Messenger {
            fun send(message: String): String {
                return "default:" + message
            }
        }

        class Push : Messenger {
            override fun send(message: String): String {
                return "push:" + message
            }
        }

        fun main() {
            val base: Messenger = Push()
            println(base.send("ok"))
        }
    "#,
    );
    assert_eq!(out, &["push:ok"]);
}

#[test]
fn test_interface_property_requirements() {
    let out = run_prints(
        r#"
        interface Identifiable {
            val id: Int
        }

        class Record(override val id: Int) : Identifiable

        fun main() {
            val r: Identifiable = Record(12)
            println(r.id)
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_interface_with_implementation_chain() {
    let out = run_prints(
        r#"
        interface Speaker {
            fun speak(): String
        }

        interface LoudSpeaker : Speaker {
            override fun speak(): String {
                return "loud"
            }
        }

        class Alarm : LoudSpeaker

        fun main() {
            val a = Alarm()
            println(a.speak())
        }
    "#,
    );
    assert_eq!(out, &["loud"]);
}

#[test]
fn test_interface_typed_reference_calls_override() {
    let out = run_prints(
        r#"
        interface Worker {
            fun work(): Int
        }

        class Engineer : Worker {
            override fun work(): Int = 3
        }

        fun report(w: Worker): Int {
            return w.work()
        }

        fun main() {
            println(report(Engineer()))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_interface_default_to_override() {
    let out = run_prints(
        r#"
        interface Source {
            fun value(): Int {
                return 1
            }
        }

        class Provider : Source {
            override fun value(): Int {
                return 2
            }
        }

        fun main() {
            val s: Source = Provider()
            println(s.value())
        }
    "#,
    );
    assert_eq!(out, &["2", "2"]);
}

#[test]
fn test_interface_mixed_implementation() {
    let out = run_prints(
        r#"
        interface A {
            fun name(): String
        }

        interface B {
            fun count(): Int
        }

        class Combo(val label: String, val amount: Int) : A, B {
            override fun name(): String = label
            override fun count(): Int = amount
        }

        fun main() {
            val c = Combo("item", 4)
            println(c.name())
            println(c.count())
        }
    "#,
    );
    assert_eq!(out, &["item", "4"]);
}

#[test]
fn test_interface_object_reference_dispatch() {
    let out = run_prints(
        r#"
        interface Status {
            fun code(): Int
        }

        class Offline : Status {
            override fun code(): Int = 0
        }

        class Online : Status {
            override fun code(): Int = 1
        }

        fun describe(status: Status): String {
            return if (status.code() == 0) "off" else "on"
        }

        fun main() {
            println(describe(Offline()))
            println(describe(Online()))
        }
    "#,
    );
    assert_eq!(out, &["off", "on"]);
}

#[test]
fn test_interface_default_property() {
    let out = run_prints(
        r#"
interface Identity { val id: Int }; class Item(override val id: Int): Identity; fun main() { println(Item(7).id) }
"#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_interface_multiple_defaults() {
    let out = run_prints(
        r#"
interface A { fun x(): String = "a" }; interface B { fun y(): String = "b" }; class C : A, B; fun main() { val b: B = C(); println(b.y()) }
"#,
    );
    assert_eq!(out, &["b"]);
}

#[test]
fn test_interface_inheritance_chain() {
    let out = run_prints(
        r#"
interface Root { fun name(): String }; interface Mid : Root { fun suffix(): String = ".mid" }; interface Leaf : Mid { fun tail(): String = ".leaf" }; class Thing : Leaf { override fun name(): String = "t" }; fun main() { val t = Thing(); println(t.name() + t.suffix() + t.tail()) }
"#,
    );
    assert_eq!(out, &["t.mid.leaf"]);
}

#[test]
fn test_interface_property_override() {
    let out = run_prints(
        r#"
interface Readable { val value: String }; class A : Readable { override val value = "alpha" }; fun main() { val r: Readable = A(); println(r.value) }
"#,
    );
    assert_eq!(out, &["alpha"]);
}

#[test]
fn test_interface_as_function_argument() {
    let out = run_prints(
        r#"
interface Caller { fun call(): Int }; class Num : Caller { override fun call(): Int = 8 }; fun invoke(c: Caller) = c.call(); fun main() { println(invoke(Num())) }
"#,
    );
    assert_eq!(out, &["8"]);
}

#[test]
fn test_interface_casting_behavior() {
    let out = run_prints(
        r#"
interface X { fun value(): Int }; class Y: X { override fun value(): Int = 4 }; class Z : X { override fun value(): Int = 9 }; fun main() { val x: X = Y(); println(x.value() + Z().value()) }
"#,
    );
    assert_eq!(out, &["13"]);
}

#[test]
fn test_interface_reflection_call() {
    let out = run_prints(
        r#"
interface Noting { fun tick(): String }; class Engine : Noting { override fun tick(): String = "ok" }; fun log(n: Noting) { println(n.tick()) }; fun main() { log(Engine()) }
"#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_interface_with_data_like_impl() {
    let out = run_prints(
        r#"
interface Shape { fun area(): Int }; class Square(val side: Int) : Shape { override fun area(): Int = side * side }; fun main() { val s: Shape = Square(6); println(s.area()) }
"#,
    );
    assert_eq!(out, &["36"]);
}

#[test]
fn test_interface_default_and_override() {
    let out = run_prints(
        r#"
interface Source { fun text(): String = "base" }; class OverrideSource : Source { override fun text(): String = "child" }; fun main() { val s: Source = OverrideSource(); println(s.text()) }
"#,
    );
    assert_eq!(out, &["child"]);
}

#[test]
fn test_interface_with_property_state() {
    let out = run_prints(
        r#"
interface Counted { var count: Int }; class Counter: Counted { override var count: Int = 0 }; fun main() { val c: Counted = Counter(); c.count = 3; println(c.count) }
"#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_interface_static_lookup_like() {
    let out = run_prints(
        r#"
interface Id { fun id(): Int }; class One : Id { override fun id(): Int = 1 }; class Two : Id { override fun id(): Int = 2 }; fun main() { val list: Array<Id> = arrayOf(One(), Two()); println(list[0].id() + list[1].id()) }
"#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_interface_default_message() {
    let out = run_prints(
        r#"
interface Message { fun msg(): String = "hi" }; class Holder: Message; fun main() { val value: Message = Holder(); println(value.msg()) }
"#,
    );
    assert_eq!(out, &["hi"]);
}

#[test]
fn test_interface_boundaries() {
    let out = run_prints(
        r#"
interface Counter { fun value(): Int }; class A : Counter { override fun value(): Int = 1 }; class B : Counter { override fun value(): Int = 2 }; fun main() { println(A().value()); println(B().value()) }
"#,
    );
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_interface_object_dispatch() {
    let out = run_prints(
        r#"
interface Op { fun run(x: Int): Int }; class Add : Op { override fun run(x: Int): Int = x + 1 }; class Mul : Op { override fun run(x: Int): Int = x * 2 }; fun apply(op: Op, value: Int): Int = op.run(value); fun main() { println(apply(Add(), 3)); println(apply(Mul(), 3)) }
"#,
    );
    assert_eq!(out, &["4", "6"]);
}

#[test]
fn test_interface_maybe_aliasing() {
    let out = run_prints(
        r#"
interface Named { fun name(): String }; class L : Named { override fun name(): String = "lab" }; fun main() { val n: Named = L(); val m: Named = n; println(m.name()) }
"#,
    );
    assert_eq!(out, &["lab"]);
}

#[test]
fn test_interface_in_while_condition() {
    let out = run_prints(
        r#"
interface Marker { fun hit(): Boolean }; class Yes: Marker { override fun hit(): Boolean = true }; class No: Marker { override fun hit(): Boolean = false }; fun main() { var score = 0; for (m in arrayOf(Yes(), No())) { if (m.hit()) score += 1 }; println(score) }
"#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_interface_with_returning_impl() {
    let out = run_prints(
        r#"
interface Factory { fun make(): Int }; class Maker : Factory { override fun make(): Int = 10 }; fun makeSomething(factory: Factory): Int = factory.make(); fun main() { println(makeSomething(Maker())) }
"#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_interface_nested_implementation() {
    let out = run_prints(
        r#"
interface Reader { fun read(): String }; class Bundle { fun create(): Reader = object : Reader { override fun read(): String = "yes" } }; fun main() { val b = Bundle(); println(b.create().read()) }
"#,
    );
    assert_eq!(out, &["yes"]);
}

#[test]
fn test_interface_property_polymorphism() {
    let out = run_prints(
        r#"
interface Flag { val code: Int }; class True: Flag { override val code = 1 }; class False: Flag { override val code = 0 }; fun main() { val f: Flag = True(); val g: Flag = False(); println(f.code + g.code) }
"#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_interface_default_method_conflict_and_explicit_super() {
    let out = run_prints(
        r#"
        interface Left {
            fun label(): String = "left"
        }

        interface Right {
            fun label(): String = "right"
        }

        class Both : Left, Right {
            override fun label(): String = super<Left>.label()
        }

        fun main() {
            val both: Left = Both()
            println(both.label())
            println((both as Right).label())
        }
    "#,
    );
    assert_eq!(out, &["left", "left"]);
}

#[test]
fn test_interface_mutable_property_backing_access() {
    let out = run_prints(
        r#"
        interface Counter {
            var value: Int
        }

        class Store(initial: Int) : Counter {
            override var value: Int = initial
        }

        fun main() {
            val c: Counter = Store(2)
            println(c.value)
            c.value += 5
            println(c.value)
        }
    "#,
    );
    assert_eq!(out, &["2", "7"]);
}

#[test]
fn test_interface_generic_method_bound_by_implementer() {
    let out = run_prints(
        r#"
        interface Formatter {
            fun <T : Number> format(value: T): String
        }

        class IntFormatter : Formatter {
            override fun <T : Number> format(value: T): String = "n:" + value.toInt().toString()
        }

        fun main() {
            val f: Formatter = IntFormatter()
            println(f.format(12))
            println(f.format(12.4))
        }
    "#,
    );
    assert_eq!(out, &["n:12", "n:12"]);
}

#[test]
fn test_interface_anonymous_object_with_capture() {
    let out = run_prints(
        r#"
        interface Supplier {
            fun value(): String
        }

        fun main() {
            val prefix = "hello "
            val supplier = object : Supplier {
                override fun value(): String = prefix + "world"
            }
            println(supplier.value())
        }
    "#,
    );
    assert_eq!(out, &["hello world"]);
}

#[test]
fn test_interface_nullable_receiver_and_safe_call() {
    let out = run_prints(
        r#"
        interface Reporter {
            fun report(): String
        }

        class Logger : Reporter {
            override fun report(): String = "ok"
        }

        fun main() {
            val good: Reporter? = Logger()
            val bad: Reporter? = null
            println(good?.report() ?: "missing")
            println(bad?.report() ?: "missing")
        }
    "#,
    );
    assert_eq!(out, &["ok", "missing"]);
}

#[test]
fn test_interface_extension_with_override_chain() {
    let out = run_prints(
        r#"
        interface Parent {
            fun base(): String = "base"
        }

        interface Child : Parent {
            override fun base(): String = "child"
        }

        class Impl : Child

        fun Parent.tag(): String = this.base() + ":tagged"

        fun main() {
            val c: Child = Impl()
            println(c.tag())
            println((c as Parent).base())
            println(c.base())
        }
    "#,
    );
    assert_eq!(out, &["child:tagged", "child", "child"]);
}

#[test]
fn test_interface_array_dispatch_across_types() {
    let out = run_prints(
        r#"
        interface Token {
            fun kind(): String
        }

        class Alpha : Token {
            override fun kind(): String = "alpha"
        }

        class Beta : Token {
            override fun kind(): String = "beta"
        }

        fun main() {
            val tokens: Array<Token> = arrayOf(Alpha(), Beta(), Alpha())
            var alpha = 0
            var beta = 0
            for (token in tokens) {
                when (token.kind()) {
                    "alpha" -> alpha += 1
                    "beta" -> beta += 1
                }
            }
            println(alpha)
            println(beta)
        }
    "#,
    );
    assert_eq!(out, &["2", "1"]);
}

#[test]
fn test_interface_default_method_conflict_override() {
    let out = run_prints(
        r#"
        interface A {
            fun label(): String = "A"
        }
        interface B {
            fun label(): String = "B"
        }
        class C : A, B {
            override fun label(): String = super<A>.label() + "+" + super<B>.label()
        }
        fun main() {
            val value: A = C()
            println(value.label())
        }
    "#,
    );
    assert_eq!(out, &["A+B"]);
}

#[test]
fn test_interface_property_backing_state() {
    let out = run_prints(
        r#"
        interface Counter {
            var count: Int
        }

        class Stateful : Counter {
            private var backing = 1
            override var count: Int
                get() = backing
                set(value) { backing = value }
        }

        fun main() {
            val c: Counter = Stateful()
            println(c.count)
            c.count += 4
            println(c.count)
        }
    "#,
    );
    assert_eq!(out, &["1", "5"]);
}

#[test]
fn test_interface_generic_contract() {
    let out = run_prints(
        r#"
        interface Boxed<T> {
            val payload: T
            fun unwrap(): T
        }

        class IntBox(override val payload: Int) : Boxed<Int> {
            override fun unwrap(): Int = payload
        }

        fun main() {
            val value: Boxed<Int> = IntBox(9)
            println(value.unwrap())
            println(value.payload)
        }
    "#,
    );
    assert_eq!(out, &["9", "9"]);
}

#[test]
fn test_interface_generic_constraint_in_function() {
    let out = run_prints(
        r#"
        interface Convertible<T> {
            fun convert(): T
        }

        class WrapInt : Convertible<String> {
            override fun convert(): String = "x"
        }

        fun <T> render(value: Convertible<T>): String {
            return value.convert().toString()
        }

        fun main() {
            println(render(WrapInt()))
        }
    "#,
    );
    assert_eq!(out, &["x"]);
}

#[test]
fn test_interface_implements_in_local_scope() {
    let out = run_prints(
        r#"
        interface Callable {
            fun call(): Int
        }

        fun main() {
            class Local : Callable {
                override fun call(): Int = 4
            }
            val item = Local()
            val c: Callable = item
            println(c.call())
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_interface_object_expression_mutating_capture() {
    let out = run_prints(
        r#"
        interface Mutator {
            fun next(): Int
        }

        fun main() {
            var n = 1
            val m: Mutator = object : Mutator {
                override fun next(): Int {
                    val out = n
                    n += 1
                    return out
                }
            }
            println(m.next())
            println(m.next())
        }
    "#,
    );
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_interface_array_dispatch() {
    let out = run_prints(
        r#"
        interface Score { fun score(): Int }
        class A : Score { override fun score(): Int = 1 }
        class B : Score { override fun score(): Int = 2 }
        class C : Score { override fun score(): Int = 3 }

        fun total(items: Array<Score>): Int {
            var sum = 0
            for (item in items) {
                sum += item.score()
            }
            return sum
        }

        fun main() {
            println(total(arrayOf(A(), B(), C())))
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_interface_extension_dispatch() {
    let out = run_prints(
        r##"
        interface Taggable {
            fun tag(): String
        }

        class Item : Taggable {
            override fun tag(): String = "item"
        }

        fun Taggable.label(): String {
            return this.tag() + "#"
        }

        fun main() {
            val value: Taggable = Item()
            println(value.label())
        }
    "##,
    );
    assert_eq!(out, &["item#"]);
}

#[test]
fn test_interface_null_cast_to_nullable() {
    let out = run_prints(
        r#"
        interface Reader { fun read(): Int }

        class NumberReader : Reader {
            override fun read(): Int = 7
        }

        fun main() {
            val source: Any? = null
            val value = source as? Reader
            println(value == null)
            println((NumberReader() as Reader).read())
        }
    "#,
    );
    assert_eq!(out, &["true", "7"]);
}

#[test]
fn test_interface_casting_wrong_type_falls_back() {
    let out = run_prints(
        r#"
        interface First { fun id(): Int }
        interface Second { fun mark(): String }

        class FirstImpl : First { override fun id(): Int = 5 }

        fun main() {
            val value: First = FirstImpl()
            val first = value as First
            val second = value as? Second
            println(first.id())
            println(second == null)
        }
    "#,
    );
    assert_eq!(out, &["5", "true"]);
}

#[test]
fn test_interface_inheritance_property_override_chain() {
    let out = run_prints(
        r#"
        interface Read {
            val source: String
        }

        interface Cache : Read {
            override val source: String
            fun open(): String = "cached:" + source
        }

        class SourceFile : Cache {
            override val source: String = "in-memory"
        }

        fun main() {
            val item: Cache = SourceFile()
            println(item.source)
            println(item.open())
        }
    "#,
    );
    assert_eq!(out, &["in-memory", "cached:in-memory"]);
}

#[test]
fn test_interface_default_used_for_abstract_missing_impl() {
    let out = run_prints(
        r#"
        interface Protocol {
            fun route(): String = "default-route"
            fun name(): String
        }

        class Service : Protocol {
            override fun name(): String = "svc"
        }

        class OverrideService : Protocol {
            override fun route(): String = "custom-route"
            override fun name(): String = "custom-svc"
        }

        fun main() {
            val base: Protocol = Service()
            val overrideSvc: Protocol = OverrideService()
            println(base.route())
            println(base.name())
            println(overrideSvc.route())
        }
    "#,
    );
    assert_eq!(out, &["default-route", "svc", "custom-route"]);
}
