use crate::helpers::run_prints;

#[test]
fn test_class_inheritance_and_override() {
    let out = run_prints(r#"
        open class Animal(val name: String) {
            open fun speak() {
                println(name + " makes a sound")
            }
        }

        class Dog(name: String) : Animal(name) {
            override fun speak() {
                println(name + " barks")
            }
        }

        fun main() {
            val dog = Dog("Rex")
            dog.speak()
        }
    "#);
    assert_eq!(out, &["Rex barks"]);
}

#[test]
fn test_companion_object() {
    let out = run_prints(r#"
        class Factory {
            companion object {
                fun create(): String {
                    return "Instance Created"
                }
            }
        }

        fun main() {
            println(Factory.create())
        }
    "#);
    assert_eq!(out, &["Instance Created"]);
}

#[test]
fn test_default_parameters() {
    let out = run_prints(r#"
        fun greet(name: String = "World") {
            println("Hello " + name)
        }

        fun main() {
            greet()
            greet("Kotlin")
        }
    "#);
    assert_eq!(out, &["Hello World", "Hello Kotlin"]);
}

#[test]
fn test_when_expression_branching() {
    let out = run_prints(r#"
        fun evaluate(x: Int) {
            when (x) {
                1 -> println("one")
                2 -> println("two")
                else -> println("other")
            }
        }

        fun main() {
            evaluate(1)
            evaluate(2)
            evaluate(99)
        }
    "#);
    assert_eq!(out, &["one", "two", "other"]);
}

#[test]
fn test_abstract_class_and_subclass() {
    let out = run_prints(r#"
        abstract class Shape {
            abstract fun area(): Int
            fun describe() {
                println("Shape area is " + area())
            }
        }

        class Square(val side: Int) : Shape() {
            override fun area(): Int = side * side
        }

        fun main() {
            val s = Square(5)
            s.describe()
        }
    "#);
    assert_eq!(out, &["Shape area is 25"]);
}

#[test]
fn test_multilevel_inheritance() {
    let out = run_prints(r#"
        open class Vehicle(val speed: Int)
        open class Car(speed: Int, val brand: String) : Vehicle(speed)
        class SportsCar(speed: Int, brand: String) : Car(speed, brand)

        fun main() {
            val sc = SportsCar(250, "Ferrari")
            println(sc.brand)
            println(sc.speed)
        }
    "#);
    assert_eq!(out, &["Ferrari", "250"]);
}

#[test]
fn test_when_multiple_cases() {
    let out = run_prints(r#"
        fun main() {
            val day = 6
            when (day) {
                1, 2, 3, 4, 5 -> println("weekday")
                6, 7 -> println("weekend")
            }
        }
    "#);
    assert_eq!(out, &["weekend"]);
}

#[test]
fn test_companion_object_method() {
    let out = run_prints(r#"
        class Counter {
            companion object {
                fun getInit(): Int = 10
            }
        }

        fun main() {
            println(Counter.getInit())
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_extension_function() {
    let out = run_prints(r#"
        class Text(val value: String)

        fun Text.emphasize(): String {
            return value + value
        }

        fun main() {
            val text = Text("hello")
            println(text.emphasize())
        }
    "#);
    assert_eq!(out, &["hellohello"]);
}

#[test]
fn test_sealed_hierarchy() {
    let out = run_prints(r#"
        sealed class Result {
            class Ok(val value: Int) : Result()
            class Error(val message: String) : Result()
        }

        fun format(result: Result): String {
            return when (result) {
                is Result.Ok -> "ok:" + (result.value)
                is Result.Error -> "error:" + (result.message)
            }
        }

        fun main() {
            val good = Result.Ok(7)
            val bad = Result.Error("bad")
            println(format(good))
            println(format(bad))
        }
    "#);
    assert_eq!(out, &["ok:7", "error:bad"]);
}

#[test]
fn test_data_class_style_constructor_shape() {
    let out = run_prints(r#"
        data class User(val name: String, val age: Int)

        fun main() {
            val first = User("Ada", 25)
            println(first.name)
            println(first.age)
        }
    "#);
    assert_eq!(out, &["Ada", "25"]);
}

#[test]
fn test_advanced_extension_with_nullable_receiver() {
    let out = run_prints(r#"
        class Holder(val value: Int)

        fun Holder.incremented(): Holder {
            return Holder(this.value + 1)
        }

        fun main() {
            val h = Holder(10)
            println(h.incremented().value)
        }
    "#);
    assert_eq!(out, &["11"]);
}

#[test]
fn test_advanced_sealed_with_deeper_matching() {
    let out = run_prints(r#"
        sealed class Status {
            class Ok(val message: String) : Status()
            class Error(val code: Int) : Status()
        }

        fun summarize(s: Status): String {
            return when (s) {
                is Status.Ok -> s.message
                is Status.Error -> "E" + s.code
            }
        }

        fun main() {
            println(summarize(Status.Ok("fine")))
            println(summarize(Status.Error(42)))
        }
    "#);
    assert_eq!(out, &["fine", "E42"]);
}

#[test]
fn test_advanced_when_with_in_conditions() {
    let out = run_prints(r#"
        fun main() {
            val score = 77
            val label = when (score) {
                in 90..100 -> "A"
                in 80..89 -> "B"
                in 70..79 -> "C"
                else -> "F"
            }
            println(label)
        }
    "#);
    assert_eq!(out, &["C"]);
}

#[test]
fn test_advanced_try_expression_return() {
    let out = run_prints(r#"
        fun compute(): Int {
            return try {
                20 / 2
            } catch (e: Exception) {
                0
            }
        }

        fun main() {
            println(compute())
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_advanced_inheritance_chain() {
    let out = run_prints(r#"
        open class Base {
            open fun id(): String {
                return "base"
            }
        }

        open class Mid : Base() {
            override fun id(): String {
                return "mid"
            }
        }

        class Leaf : Mid() {
            override fun id(): String {
                return super.id() + "+leaf"
            }
        }

        fun main() {
            val l = Leaf()
            println(l.id())
        }
    "#);
    assert_eq!(out, &["mid+leaf"]);
}

#[test]
fn test_advanced_data_class_roundtrip() {
    let out = run_prints(r#"
        data class User(val name: String, val age: Int)

        fun main() {
            val first = User("ivy", 11)
            val second = User("ivy", 11)
            println(first == second)
            println(first.name)
            println(first.age)
        }
    "#);
    assert_eq!(out, &["true", "ivy", "11"]);
}

#[test]
fn test_advanced_sealed_default() {
    let out = run_prints(r#"
sealed class Node { class A : Node(); class B : Node() }; fun main() { val n: Node = Node.A(); if (n is Node.A) { println("a") } }
"#);
    assert_eq!(out, &["a"]);
}

#[test]
fn test_advanced_when_without_subject() {
    let out = run_prints(r#"
fun main() { val x = 4; val y = if (x < 0) 1 else 2; println(y) }
"#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_advanced_inheritance_chain_override() {
    let out = run_prints(r#"
open class A { open fun value(): Int = 1 }; open class B : A() { override fun value(): Int = 2 }; class C : B() { override fun value(): Int = 3 }; fun main() { println(C().value()) }
"#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_advanced_data_copy() {
    let out = run_prints(r#"
data class Pair(val a: Int, val b: Int); fun main() { val p = Pair(1, 2); val q = p.copy(b = 3); println(q.a); println(q.b) }
"#);
    assert_eq!(out, &["1", "3"]);
}

#[test]
fn test_advanced_extension_in_advanced() {
    let out = run_prints(r#"
class Holder(val value: Int); fun Holder.double() = value * 2; fun main() { println(Holder(4).double()) }
"#);
    assert_eq!(out, &["8"]);
}

#[test]
fn test_advanced_object_expression_chain() {
    let out = run_prints(r#"
interface Flag { fun value(): String }; fun make() = object : Flag { override fun value(): String = "go" }; fun main() { println(make().value()) }
"#);
    assert_eq!(out, &["go"]);
}

#[test]
fn test_advanced_generic_like() {
    let out = run_prints(r#"
open class Box(val item: Int) { fun value(): Int = item }; fun main() { val b = Box(7); println(b.value()) }
"#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_advanced_try_with_when() {
    let out = run_prints(r#"
fun main() { try { println("ok") } catch (e: Exception) { println("bad") } finally { println("end") } }
"#);
    assert_eq!(out, &["ok", "end"]);
}

#[test]
fn test_advanced_looped_while_for() {
    let out = run_prints(r#"
fun main() { var total = 0; for (i in 1..3) { total += i }; var x = 0; while (x < 2) { total += 1; x += 1 }; println(total) }
"#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_advanced_elvis_in_advanced() {
    let out = run_prints(r#"
fun main() { val text: String? = null; println(text ?: "none") }
"#);
    assert_eq!(out, &["none"]);
}

#[test]
fn test_advanced_cast_in_advanced() {
    let out = run_prints(r#"
fun main() { val value: Any = 2; println(value is Int); println((value as Int) + 3) }
"#);
    assert_eq!(out, &["true", "5"]);
}

#[test]
fn test_advanced_multiple_interfaces() {
    let out = run_prints(r#"
interface A { fun a(): String }; interface B { fun b(): String }; class C : A, B { override fun a() = "a"; override fun b() = "b" }; fun main() { val c: A = C(); val d: B = C(); println(c.a()); println(d.b()) }
"#);
    assert_eq!(out, &["a", "b"]);
}

#[test]
fn test_advanced_nested_conditional() {
    let out = run_prints(r#"
fun score(x: Int): String { return if (x > 10) "high" else if (x > 5) "mid" else "low" }; fun main() { println(score(11)); println(score(3)) }
"#);
    assert_eq!(out, &["high", "low"]);
}

#[test]
fn test_advanced_when_subject_evaluated_once() {
    let out = run_prints(r#"
        var calls = 0

        fun tapped(): Int {
            calls += 1
            return calls
        }

        fun main() {
            val status = when (tapped()) {
                1 -> "first"
                2 -> "second"
                else -> "other"
            }
            println(status)
            println(calls)
        }
    "#);
    assert_eq!(out, &["first", "1"]);
}

#[test]
fn test_advanced_override_open_property() {
    let out = run_prints(r#"
        open class Vehicle {
            open val kind: String = "vehicle"
        }

        class Car : Vehicle() {
            override val kind: String = "car"
        }

        fun main() {
            val v: Vehicle = Car()
            println(v.kind)
        }
    "#);
    assert_eq!(out, &["car"]);
}

#[test]
fn test_advanced_companion_object_state() {
    let out = run_prints(r#"
        class Counter {
            companion object {
                var hits: Int = 0

                fun next(): Int {
                    hits += 1
                    return hits
                }
            }
        }

        fun main() {
            println(Counter.next())
            println(Counter.next())
            println(Counter.hits)
        }
    "#);
    assert_eq!(out, &["1", "2", "2"]);
}

#[test]
fn test_advanced_nested_object_expression_with_state() {
    let out = run_prints(r#"
        fun makeCounter() = object {
            var value: Int = 0

            fun inc(): Int {
                value += 1
                return value
            }
        }

        fun main() {
            val c = makeCounter()
            println(c.inc())
            println(c.inc())
        }
    "#);
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_advanced_data_class_destructure_and_mutate_copy() {
    let out = run_prints(r#"
        data class Pair(val left: Int, val right: Int)

        fun main() {
            val original = Pair(10, 20)
            val (a, b) = original
            val updated = original.copy(right = 99)
            println(a)
            println(b)
            println(updated.left)
            println(updated.right)
        }
    "#);
    assert_eq!(out, &["10", "20", "10", "99"]);
}

#[test]
fn test_advanced_try_finally_with_return_paths() {
    let out = run_prints(r#"
        var marker = ""

        fun evaluate(use_fast: Boolean): Int {
            try {
                if (use_fast) {
                    return 7
                } else {
                    return 11
                }
            } finally {
                marker += "f"
            }
        }

        fun main() {
            println(evaluate(true))
            println(evaluate(false))
            println(marker)
        }
    "#);
    assert_eq!(out, &["7", "11", "ff"]);
}

#[test]
fn test_advanced_when_without_subject_guarded() {
    let out = run_prints(r#"
        fun main() {
            val x = 0
            val label = when {
                x > 0 -> "positive"
                x == 0 -> "zero"
                else -> "negative"
            }
            println(label)
        }
    "#);
    assert_eq!(out, &["zero"]);
}
