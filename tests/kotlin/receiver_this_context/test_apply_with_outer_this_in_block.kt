// vybe-test: kotlin/receiver_this_context/test_apply_with_outer_this_in_block
// origin: languages/kotlin/tests/kotlin/test_receiver_this_context.rs

class Profile(val base: String) {
            val id: String = "x"
            fun build(): String =
                StringBuilder().apply {
                    this@Profile.base.let { append(it) }
                    append(":")
                    append(id)
                }.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Profile("p").build()).toString(), "p:x")
        }
