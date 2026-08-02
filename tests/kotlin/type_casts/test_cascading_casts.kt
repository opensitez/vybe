// vybe-test: kotlin/type_casts/test_cascading_casts
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

open class Base
class Child : Base()
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val any: Base = Child()
val child = any as Child
__check((child is Child).toString(), "true") }
