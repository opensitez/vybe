// vybe-test: kotlin/scoping_functions/test_also_chain_keeps_reference
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val log = mutableListOf<String>()
            val values = mutableListOf(10)
                .also { log.add("initial-" + it.size.toString()) }
                .also { it.add(20) }
                .also { log.add("after-" + it.size.toString()) }
            __check((values.joinToString(";")).toString(), "10;20")
            __check((log.joinToString(",")).toString(), "initial-1,after-2")
        }
