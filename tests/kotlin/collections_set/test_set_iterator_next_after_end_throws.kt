// vybe-test: kotlin/collections_set/test_set_iterator_next_after_end_throws
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1)
            val it = values.iterator()
            __check((it.next()).toString(), "1")
            try {
                it.next()
            } catch (e: NoSuchElementException) {
                __check(("done").toString(), "done")
            }
        }
