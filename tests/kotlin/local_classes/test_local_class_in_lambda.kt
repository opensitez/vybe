// vybe-test: kotlin/local_classes/test_local_class_in_lambda
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = { value: Int ->
                class Local(val doubled: Int)
                Local(value * 2)
            }
            __check((v(3).doubled).toString(), "6")
        }
