// vybe-test: kotlin/annotations/test_multiple_custom_annotations_on_member
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

annotation class One
        annotation class Two

        @One
        @Two
        fun both(): Int {
            return 2 + 3
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((both()).toString(), "5")
        }
