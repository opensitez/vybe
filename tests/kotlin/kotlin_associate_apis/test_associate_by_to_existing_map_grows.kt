// vybe-test: kotlin/kotlin_associate_apis/test_associate_by_to_existing_map_grows
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = mutableMapOf<Int, String>()
            out[99] = "seed"
            listOf("a", "bb").associateByTo(out, { it.length }, { it })
            __check((out[99]).toString(), "seed")
            __check((out[1]).toString(), "a")
            __check((out[2]).toString(), "bb")
        }
