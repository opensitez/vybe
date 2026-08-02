// vybe-test: kotlin/class_delegation/test_delegate_object_expression
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Messenger { fun message(): String }

        class Proxy(delegate: Messenger) : Messenger by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Proxy(object : Messenger {
                override fun message() = "from object"
            })
            __check((p.message()).toString(), "from object")
        }
