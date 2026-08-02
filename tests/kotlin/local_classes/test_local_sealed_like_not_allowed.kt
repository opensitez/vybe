// vybe-test: kotlin/local_classes/test_local_sealed_like_not_allowed
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            open class Local(val v: Int)
            class Derived : Local(2)
            __check((Derived().v).toString(), "2")
        }
