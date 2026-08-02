// vybe-test: kotlin/local_classes/test_local_class_with_extension_fn
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Word(val value: String)
            fun Word.quoted() = "\"$value\""
            __check((Word("k").quoted()).toString(), "\"k\"")
        }
