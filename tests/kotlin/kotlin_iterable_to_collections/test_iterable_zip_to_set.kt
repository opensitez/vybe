// vybe-test: kotlin/kotlin_iterable_to_collections/test_iterable_zip_to_set
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val zipped = (1..3).zip("abc")
            val s = zipped.toSet()
            __check((s.size).toString(), "3")
            __check((s.any { it.first == 2 }.toString()).toString(), "true")
        }
