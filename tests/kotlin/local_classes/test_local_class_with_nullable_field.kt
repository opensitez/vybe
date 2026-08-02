// vybe-test: kotlin/local_classes/test_local_class_with_nullable_field
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class L(val v: String?)
            __check((L(null).v).toString(), "null")
            __check((L("x").v).toString(), "x")
        }
