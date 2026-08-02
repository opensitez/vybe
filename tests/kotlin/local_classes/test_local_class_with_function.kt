// vybe-test: kotlin/local_classes/test_local_class_with_function
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class C {
                fun id(v: Int): Int = v + 1
            }
            __check((C().id(4)).toString(), "5")
        }
