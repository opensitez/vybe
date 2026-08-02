// vybe-test: kotlin/iterator_protocol/test_iterator_map_side_effect_order
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mutableListOf(1, 2, 3)
            var log = ""
            val it = source.map {
                log += "#" + it
                it
            }.iterator()
            __check((it.next()).toString(), "1")
            __check((it.next()).toString(), "2")
            __check((log).toString(), "#1#2#3")
        }
