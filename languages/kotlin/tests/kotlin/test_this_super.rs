kotlin_run_test!(
    test_this_self_reference,
    r#"fun main() { class K { fun id(): String = this.toString() }; println(K().id().isNotEmpty()) }"#,
    &["true"]
);

kotlin_run_test!(
    test_this_inside_member_access,
    r#"class K { val v = 1; fun out() = this.v + 1 }
fun main() { println(K().out()) }"#,
    &["2"]
);

kotlin_run_test!(
    test_this_in_constructor,
    r#"class K(val x: Int) { init { println(this.x) } }
fun main() { K(3) }"#,
    &["3"]
);

kotlin_run_test!(
    test_this_in_inner_class,
    r#"class Outer {
        val tag = "outer"
        inner class Inner { fun outerTag() = this@Outer.tag }
    }
    fun main() { println(Outer().Inner().outerTag()) }"#,
    &["outer"]
);

kotlin_run_test!(
    test_this_in_nested_lambda,
    r#"class K { fun call(): String { val f = { this.toString() }; return f() } }
fun main() { println(K().call().isNotEmpty()) }"#,
    &["true"]
);

kotlin_run_test!(
    test_super_simple_override,
    r#"open class A { open fun tag() = "a" }
class B: A() { override fun tag() = super.tag() + "+b" }
fun main() { println(B().tag()) }"#,
    &["a+b"]
);

kotlin_run_test!(
    test_super_property_reference,
    r#"open class A { open val v = 1 }
class B: A() { override val v = super.v + 2 }
fun main() { println(B().v) }"#,
    &["3"]
);

kotlin_run_test!(
    test_super_function_in_interface,
    r#"interface I { fun tag() = "i" }
open class A { open fun tag() = "a" }
class B: A(), I { override fun tag() = super<A>.tag() + super<I>.tag() }"#,
    &["ai"]
);

kotlin_run_test!(
    test_this_in_inheritance_chain,
    r#"open class A { open fun who() = "A" }
open class B: A() { override fun who() = "B" }
class C: B() { override fun who() = super<B>.who() + "C" }
fun main() { println(C().who()) }"#,
    &["BC"]
);

kotlin_run_test!(
    test_super_class_in_inner,
    r#"open class A { open fun label() = "base" }
class B : A() {
    inner class Inner { fun label() = super@B.label() }
}
fun main() { println(B().Inner().label()) }"#,
    &["base"]
);

kotlin_run_test!(
    test_this_outer_in_outer_function,
    r#"class Outer {
        val t = "outer"
        inner class Inner { fun t() = this@Outer.t }
    }
    fun main() { println(Outer().Inner().t()) }"#,
    &["outer"]
);

kotlin_run_test!(
    test_this_outer_in_class_nesting,
    r#"class Level1 {
        val x = "one"
        class Level2(val parent: Level1) { fun read() = parent.x }
    }
    fun main() { val p = Level1(); println(Level1.Level2(p).read()) }"#,
    &["one"]
);

kotlin_run_test!(
    test_this_in_anonymous_object,
    r#"fun main() {
        val o = object {
            val value = 9
            fun read(): Int = this.value
        }
        println(o.read())
    }"#,
    &["9"]
);

kotlin_run_test!(
    test_this_in_extension_receiver,
    r#"class Box(val n: Int) { fun call() = with(this) { n + 1 } }
fun main() { println(Box(4).call()) }"#,
    &["5"]
);

kotlin_run_test!(
    test_this_in_require_context,
    r#"class Base {
        fun check() = "ok"
    }
    class Child : Base() {
        fun run() = this.check()
    }
    fun main() { println(Child().run()) }"#,
    &["ok"]
);

kotlin_run_test!(
    test_this_with_apply,
    r#"fun main() {
        val out = StringBuilder().apply { this.append("a") }.toString()
        println(out)
    }"#,
    &["a"]
);

kotlin_run_test!(
    test_this_in_run_return,
    r#"fun main() {
        val x = StringBuilder().run {
            this.append("x")
            this.toString()
        }
        println(x)
    }"#,
    &["x"]
);

kotlin_run_test!(
    test_this_in_with,
    r#"fun main() {
        val x = "k"
        val y = with(x) { this + this }
        println(y)
    }"#,
    &["kk"]
);

kotlin_run_test!(
    test_this_ref_in_to_string_override,
    r#"class A {
        override fun toString() = "A:" + this.javaClass.simpleName
    }
    fun main() { println(A().toString()) }"#,
    &["A:A"]
);

kotlin_run_test!(
    test_super_in_init_order,
    r#"open class A { open val a = "a" }
class B : A() { override val a = super.a + "+b" }
fun main() { println(B().a) }"#,
    &["a+b"]
);

kotlin_run_test!(
    test_this_reference_in_if,
    r#"class Holder {
        fun value(v: Int?): Int {
            return if (this.hashCode() > 0) (v ?: 0) + 1 else 0
        }
    }
    fun main() { println(Holder().value(3)) }"#,
    &["4"]
);

kotlin_run_test!(
    test_super_function_in_diamond,
    r#"open class A { open fun x() = "A" }
interface I { fun x() = "I" }
class B : A(), I { override fun x() = super<A>.x() }
fun main() { println(B().x()) }"#,
    &["A"]
);

kotlin_run_test!(
    test_super_interface_call,
    r#"interface I { fun y() = "I" }
class C : I { override fun y() = super<I>.y() }
fun main() { println(C().y()) }"#,
    &["I"]
);

kotlin_run_test!(
    test_this_in_secondary_constructor,
    r#"class C(val v: Int) { constructor() : this(3) { println(this.v) } }
fun main() { C() }"#,
    &["3"]
);

kotlin_run_test!(
    test_this_in_nested_object,
    r#"class K {
        private val x = 1
        fun maker() = object {
            fun value() = this@K.x
        }
    }
    fun main() { println(K().maker().value()) }"#,
    &["1"]
);

kotlin_run_test!(
    test_this_equality_identity,
    r#"class K { fun same(other: K): Boolean = this === other }
fun main() { val a = K(); val b = a; println(a.same(b)) }"#,
    &["true"]
);

kotlin_run_test!(
    test_this_is_used_in_comparison,
    r#"class K { val v = 1; fun same(other: K) = this.v == other.v }
fun main() { println(K().same(K())) }"#,
    &["true"]
);

kotlin_run_test!(
    test_super_in_accessor,
    r#"open class A { open fun get(): Int = 1 }
class B : A() { override fun get() = super.get() + 2 }
fun main() { println(B().get()) }"#,
    &["3"]
);

kotlin_run_test!(
    test_this_in_data_class_copy,
    r#"data class X(val v: Int)
fun main() {
    val x = X(1)
    val y = x.copy(v = x.v + 1)
    println(y.v)
}"#,
    &["1"]
);

kotlin_run_test!(
    test_this_in_companion_method,
    r#"class C {
        companion object { fun label(): String = "comp" }
        fun out(): String = C.label()
    }
    fun main() { println(C().out()) }"#,
    &["comp"]
);

kotlin_run_test!(
    test_super_to_string_chain,
    r#"open class A { override fun toString() = "A" }
class B : A() { override fun toString() = super.toString() + "B" }
fun main() { println(B().toString()) }"#,
    &["AB"]
);

kotlin_run_test!(
    test_this_in_when_branch,
    r#"class K { fun kind(v: Int): String = when (v) { 1 -> this.javaClass.simpleName; else -> "n" } }
fun main() { println(K().kind(1)) }"#,
    &["K"]
);

kotlin_run_test!(
    test_super_in_multilevel,
    r#"open class A { open fun depth() = "A" }
open class B : A() { override fun depth() = super<A>.depth() + "->B" }
class C : B() { override fun depth() = super<B>.depth() + "->C" }
fun main() { println(C().depth()) }"#,
    &["A->B->C"]
);

kotlin_run_test!(
    test_this_inside_try_block,
    r#"class K { fun id() = this.toString() }
fun main() { try { println(K().id().isNotEmpty()) } catch (e: Exception) { println("err") } }"#,
    &["true"]
);

kotlin_run_test!(
    test_this_in_extension_lambda,
    r#"fun K.tag() = n(this)
class K(val value: Int)
fun n(k: K) = k.value
fun main() { println(K(4).tag()) }"#,
    &["4"]
);

kotlin_run_test!(
    test_this_nested_class_calls,
    r#"class A { val v = 3
    inner class B { fun v() = this@A.v }
}
fun main() { println(A().B().v()) }"#,
    &["3"]
);

kotlin_run_test!(
    test_super_with_override_chain,
    r#"open class A { open val p = "A" }
open class B: A() { override val p = "B" }
class C: B() { override val p = super<B>.p + "C" }
fun main() { println(C().p) }"#,
    &["BC"]
);
