use crate::helpers::run_prints;

#[test]
fn test_class_declaration() {
    let out = run_prints(r#"
        class Person(val name: String, var age: Int) {
            fun greet() {
                println("I am " + name)
            }
        }

        fun main() {
            val p = Person("Alice", 30)
            p.greet()
        }
    "#);
    assert_eq!(out, &["I am Alice"]);
}

#[test]
fn test_object_declaration() {
    let out = run_prints(r#"
        object Logger {
            fun log(msg: String) {
                println("LOG: " + msg)
            }
        }

        fun main() {
            Logger.log("Started")
        }
    "#);
    assert_eq!(out, &["LOG: Started"]);
}

#[test]
fn test_class_field_mutation() {
    let out = run_prints(r#"
        class Counter(var count: Int) {
            fun inc() {
                count += 1
            }
        }

        fun main() {
            val c = Counter(10)
            c.inc()
            println(c.count)
        }
    "#);
    assert_eq!(out, &["11"]);
}

#[test]
fn test_multiple_class_instances() {
    let out = run_prints(r#"
        class Account(val id: String, var balance: Int)

        fun main() {
            val a1 = Account("A", 100)
            val a2 = Account("B", 200)
            a1.balance += 50
            println(a1.balance)
            println(a2.balance)
        }
    "#);
    assert_eq!(out, &["150", "200"]);
}

#[test]
fn test_nested_class() {
    let out = run_prints(r#"
        class Outer {
            class Nested {
                fun getMsg(): String = "nested msg"
            }
        }

        fun main() {
            val n = Outer.Nested()
            println(n.getMsg())
        }
    "#);
    assert_eq!(out, &["nested msg"]);
}

#[test]
fn test_class_multiple_methods() {
    let out = run_prints(r#"
        class Calc(val base: Int) {
            fun add(x: Int): Int = base + x
            fun mul(x: Int): Int = base * x
        }

        fun main() {
            val c = Calc(5)
            println(c.add(3))
            println(c.mul(4))
        }
    "#);
    assert_eq!(out, &["8", "20"]);
}

#[test]
fn test_class_constructor_argument() {
    let out = run_prints(r#"
        class Config(val timeout: Int)

        fun main() {
            val c = Config(30)
            println(c.timeout)
        }
    "#);
    assert_eq!(out, &["30"]);
}

#[test]
fn test_class_method_chaining() {
    let out = run_prints(r#"
        class Builder(var msg: String) {
            fun append(s: String): Builder {
                msg += s
                return this
            }
        }

        fun main() {
            val b = Builder("A")
            b.append("B").append("C")
            println(b.msg)
        }
    "#);
    assert_eq!(out, &["ABC"]);
}

#[test]
fn test_init_block_execution() {
    let out = run_prints(r#"
        class Counter {
            var count = 0
            init {
                println("init")
            }
            fun increment() {
                count += 1
            }
        }

        fun main() {
            val c = Counter()
            c.increment()
            c.increment()
            println(c.count)
        }
    "#);
    assert_eq!(out, &["init", "2"]);
}

#[test]
fn test_class_property_getter() {
    let out = run_prints(r#"
        class Box(val value: Int) {
            val doubled: Int
                get() = value * 2
        }

        fun main() {
            val box = Box(7)
            println(box.doubled)
        }
    "#);
    assert_eq!(out, &["14"]);
}

#[test]
fn test_class_inheritance_with_super() {
    let out = run_prints(r#"
        open class Parent {
            open fun name(): String = "parent"
        }

        class Child : Parent() {
            override fun name(): String {
                return super.name() + "-child"
            }
        }

        fun main() {
            val child = Child()
            println(child.name())
        }
    "#);
    assert_eq!(out, &["parent-child"]);
}

#[test]
fn test_companion_stateful_object() {
    let out = run_prints(r#"
        class Holder {
            companion object {
                var created = 0
                fun create(): Holder {
                    created += 1
                    return Holder()
                }
            }
        }

        fun main() {
            Holder.create()
            Holder.create()
            println(Holder.created)
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_class_with_init_and_field() {
    let out = run_prints(r#"
        class Meter {
            val limit: Int
            init {
                println("init")
            }
            constructor(value: Int) {
                this.limit = value
            }
            fun scale(): Int {
                return limit * 2
            }
        }

        fun main() {
            val m = Meter(7)
            println(m.scale())
        }
    "#);
    assert_eq!(out, &["init", "14"]);
}

#[test]
fn test_class_with_custom_getter_and_state() {
    let out = run_prints(r#"
        class Box {
            val base: Int = 6
            val doubled: Int
                get() = base * 2
        }

        fun main() {
            val b = Box()
            println(b.base)
            println(b.doubled)
        }
    "#);
    assert_eq!(out, &["6", "12"]);
}

#[test]
fn test_class_open_override_property() {
    let out = run_prints(r#"
        open class Base {
            open fun label(): String {
                return "base"
            }
        }

        class Child : Base() {
            override fun label(): String {
                return "child"
            }
        }

        fun main() {
            val b: Base = Child()
            println(b.label())
        }
    "#);
    assert_eq!(out, &["child"]);
}

#[test]
fn test_class_nested_and_instantiation() {
    let out = run_prints(r#"
        class Outer {
            class Inner {
                fun ping(): String {
                    return "pong"
                }
            }
        }

        fun main() {
            val i = Outer.Inner()
            println(i.ping())
        }
    "#);
    assert_eq!(out, &["pong"]);
}

#[test]
fn test_companion_object_with_var() {
    let out = run_prints(r#"
        class Counter {
            companion object {
                var created = 0
                fun track(): Int {
                    created += 1
                    return created
                }
            }
        }

        fun main() {
            println(Counter.track())
            println(Counter.track())
        }
    "#);
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_class_with_primary_default() {
    let out = run_prints(r#"
class Item(val name: String = "a") { fun mainName(): String = name }; fun main() { println(Item().mainName()) }
"#);
    assert_eq!(out, &["a"]);
}

#[test]
fn test_abstract_class_contract() {
    let out = run_prints(r#"
abstract class Node { abstract fun id(): Int }; class Leaf : Node() { override fun id(): Int = 9 }; fun main() { val n: Node = Leaf(); println(n.id()) }
"#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_private_like_state_mutation() {
    let out = run_prints(r#"
class Bank { var balance: Int = 100; fun deposit(v: Int) { balance += v }; fun withdraw(v: Int) { balance -= v }; fun total(): Int = balance }; fun main() { val b = Bank(); b.deposit(40); b.withdraw(10); println(b.total()) }
"#);
    assert_eq!(out, &["130"]);
}

#[test]
fn test_class_with_getter_only_property() {
    let out = run_prints(r#"
class Product { val price = 7; val doubled: Int get() = price * 2 }; fun main() { val p = Product(); println(p.doubled) }
"#);
    assert_eq!(out, &["14"]);
}

#[test]
fn test_class_interface_and_override_state() {
    let out = run_prints(r#"
open class Base { open fun text(): String = "base" }; class Derived: Base() { override fun text(): String = "derived" }; fun main() { val b: Base = Derived(); println(b.text()) }
"#);
    assert_eq!(out, &["derived"]);
}

#[test]
fn test_class_with_companion_counter() {
    let out = run_prints(r#"
class Factory { companion object { var index = 0; fun next(): Int { index += 1; return index } } }; fun main() { println(Factory.next()); println(Factory.next()) }
"#);
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_class_init_and_method_order() {
    let out = run_prints(r#"
class Logger { init { println("start") }; fun value(): Int = 5 }; fun main() { println(Logger().value()) }
"#);
    assert_eq!(out, &["start", "5"]);
}

#[test]
fn test_nested_class_with_reference() {
    let out = run_prints(r#"
class Outer { class Inner { fun read(): Int = 3 } }; fun main() { val i = Outer.Inner(); println(i.read()) }
"#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_class_chain_of_inheritance() {
    let out = run_prints(r#"
open class A { open fun num(): Int = 1 }; open class B : A() { override fun num(): Int = 2 }; class C : B() { override fun num(): Int = super.num() + 3 }; fun main() { println(C().num()) }
"#);
    assert_eq!(out, &["5"]);
}


#[test]
fn test_class_chain_methods() {
    let out = run_prints(r#"
        class Builder {
            var value: Int = 0
            fun set(value: Int): Builder {
                this.value = value
                return this
            }
            fun increment(step: Int): Builder {
                this.value += step
                return this
            }
            fun total(): Int {
                return this.value
            }
        }

        fun main() {
            val b = Builder()
            val result = b.set(3).increment(4).total()
            println(result)
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_class_method_overload_variation() {
    let out = run_prints(r#"
        class Math {
            fun value(x: Int): Int = x
            fun value(x: Int, y: Int): Int = x + y
        }

        fun main() {
            val m = Math()
            println(m.value(1))
            println(m.value(2, 3))
        }
    "#);
    assert_eq!(out, &["1", "5"]);
}

#[test]
fn test_class_with_constructor_sharing() {
    let out = run_prints(r#"
        class PairNode {
            val left: Int
            val right: Int

            constructor(left: Int, right: Int) {
                this.left = left
                this.right = right
            }

            constructor(value: Int) : this(value, value) {
                println("copy")
            }
        }

        fun main() {
            val p1 = PairNode(4)
            val p2 = PairNode(1, 3)
            println(p1.left)
            println(p1.right)
            println(p2.left)
            println(p2.right)
        }
    "#);
    assert_eq!(out, &["copy", "4", "4", "1", "3"]);
}

#[test]
fn test_class_with_abstract_implementation() {
    let out = run_prints(r#"
        abstract class Worker {
            abstract fun work(): Int
        }

        class Coder : Worker() {
            override fun work(): Int {
                return 9
            }
        }

        fun main() {
            val w: Worker = Coder()
            println(w.work())
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_class_secondary_constructor_delegates_to_primary() {
    let out = run_prints(r#"
        class Box {
            val value: Int
            val label: String

            constructor(value: Int) : this(value, "default") {
                println("secondary")
            }

            constructor(value: Int, label: String) {
                this.value = value
                this.label = label
            }

            fun describe(): String {
                return label + ":" + value
            }
        }

        fun main() {
            val item = Box(5)
            println(item.describe())
        }
    "#);
    assert_eq!(out, &["secondary", "default:5"]);
}

#[test]
fn test_class_property_setter_validation() {
    let out = run_prints(r#"
        class Counter {
            var value: Int = 0
                set(next) {
                    field = if (next < 0) 0 else next
                }
        }

        fun main() {
            val c = Counter()
            c.value = 5
            println(c.value)
            c.value = -3
            println(c.value)
        }
    "#);
    assert_eq!(out, &["5", "0"]);
}

#[test]
fn test_class_open_property_override_in_subclass() {
    let out = run_prints(r#"
        open class Vehicle {
            open val tag: String = "vehicle"
        }

        class Truck : Vehicle() {
            override val tag: String = "truck"
        }

        fun main() {
            val v: Vehicle = Truck()
            println(v.tag)
        }
    "#);
    assert_eq!(out, &["truck"]);
}

#[test]
fn test_class_inner_class_captures_outer_reference() {
    let out = run_prints(r#"
        class Outer(val prefix: String) {
            val marker = "!"

            inner class Inner(val value: String) {
                fun describe(): String {
                    return prefix + marker + value
                }
            }
        }

        fun main() {
            val outer = Outer("x")
            val inner = outer.Inner("y")
            println(inner.describe())
        }
    "#);
    assert_eq!(out, &["x!y"]);
}

#[test]
fn test_class_instance_init_order_with_multiple_inits() {
    let out = run_prints(r#"
        open class Base {
            init {
                println("base")
            }
            open val baseName: String = "base"
        }

        class Child : Base() {
            init {
                println("child")
            }
            override val baseName: String = "child"
        }

        fun main() {
            val c = Child()
            println(c.baseName)
        }
    "#);
    assert_eq!(out, &["base", "child", "child"]);
}

#[test]
fn test_class_data_like_projection_with_copy_behavior_not_available() {
    let out = run_prints(r#"
        class Pair(val left: Int, val right: Int)

        fun main() {
            val a = Pair(1, 2)
            val b = Pair(a.left, a.right + 1)
            println(a.left)
            println(a.right)
            println(b.right)
        }
    "#);
    assert_eq!(out, &["1", "2", "3"]);
}
