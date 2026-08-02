// vybe-test: kotlin/sealed_types/test_sealed_interface_and_data_leaves
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed interface Kind

        data class Node(val id: Int) : Kind
        class Done : Kind

        fun describe(kind: Kind): String {
            return when (kind) {
                is Node -> "node:" + kind.id.toString()
                is Done -> "done"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(Node(7))).toString(), "node:7")
            __check((describe(Done())).toString(), "done")
        }
