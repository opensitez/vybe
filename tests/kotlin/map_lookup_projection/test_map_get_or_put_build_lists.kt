// vybe-test: kotlin/map_lookup_projection/test_map_get_or_put_build_lists
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val buckets = linkedMapOf<String, MutableList<Int>>()
            buckets.getOrPut("a") { mutableListOf() }.add(1)
            buckets.getOrPut("a") { mutableListOf() }.add(2)
            buckets.getOrPut("b") { mutableListOf() }.add(9)
            __check((buckets["a"]!!.joinToString(",")).toString(), "1,2")
            __check((buckets.size).toString(), "2")
        }
