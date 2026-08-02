// vybe-test: kotlin/local_classes/test_local_class_generic_property
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Box<T>(val v: T)
            __check((Box("x").v).toString(), "x")
            __check((Box(1).v).toString(), "1")
        }
