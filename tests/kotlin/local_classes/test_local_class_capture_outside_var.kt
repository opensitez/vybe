// vybe-test: kotlin/local_classes/test_local_class_capture_outside_var
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = 10
            class Local {
                fun total(offset: Int) = base + offset
            }
            __check((Local().total(3)).toString(), "13")
        }
