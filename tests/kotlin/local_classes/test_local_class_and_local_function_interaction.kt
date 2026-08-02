// vybe-test: kotlin/local_classes/test_local_class_and_local_function_interaction
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class C(val v: Int)
            fun f(v: C): Int = v.v * 2
            __check((f(C(4))).toString(), "8")
        }
