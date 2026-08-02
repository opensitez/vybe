// vybe-test: kotlin/advanced_features/test_advanced_multiple_interfaces
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

interface A { fun a(): String }
interface B { fun b(): String }
class C : A, B { override fun a() = "a"
override fun b() = "b" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val c: A = C()
val d: B = C()
__check((c.a()).toString(), "a")
__check((d.b()).toString(), "b") }
