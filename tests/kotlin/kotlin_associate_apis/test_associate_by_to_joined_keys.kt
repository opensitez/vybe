// vybe-test: kotlin/kotlin_associate_apis/test_associate_by_to_joined_keys
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = linkedMapOf<Int, String>()
            listOf("one", "two", "ten").associateByTo(out, { it.length }, { it.first() })
            __check((out.keys.joinToString(",")).toString(), "3,")
            __check((out[3]).toString(), "t")
        }
