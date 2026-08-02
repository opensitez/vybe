// vybe-test: kotlin/local_classes/test_local_class_with_private
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Box {
                private val hidden = 5
                fun reveal() = hidden
            }
            __check((Box().reveal()).toString(), "5")
        }
