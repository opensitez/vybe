// vybe-test: kotlin/collections_set/test_to_mutable_set_roundtrip
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3)
            val mutable = values.toMutableSet()
            mutable.remove(2)
            mutable.add(4)
            __check((mutable.size).toString(), "3")
            __check((mutable.contains(2)).toString(), "false")
            __check((mutable.contains(4)).toString(), "true")
        }
