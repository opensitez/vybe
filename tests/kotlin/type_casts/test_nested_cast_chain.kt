// vybe-test: kotlin/type_casts/test_nested_cast_chain
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

open class A
class B : A()
fun castOrZero(v: A): Int { return if (v is B) 1 else 0 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((castOrZero(B())).toString(), "1")
__check((castOrZero(A())).toString(), "0") }
