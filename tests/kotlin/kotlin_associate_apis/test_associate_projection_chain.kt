// vybe-test: kotlin/kotlin_associate_apis/test_associate_projection_chain
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = listOf("aa", "bb", "ccc").associateBy { it[0] }.mapValues { it.value.length }
            __check((map.keys.joinToString(",")).toString(), "a,b,c")
            __check((map["a"]).toString(), "2")
            __check((map["c"]).toString(), "3")
        }
