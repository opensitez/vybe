// vybe-test: kotlin/interfaces/test_interface_casting_wrong_type_falls_back
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface First { fun id(): Int }
        interface Second { fun mark(): String }

        class FirstImpl : First { override fun id(): Int = 5 }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: First = FirstImpl()
            val first = value as First
            val second = value as? Second
            __check((first.id()).toString(), "5")
            __check((second == null).toString(), "true")
        }
