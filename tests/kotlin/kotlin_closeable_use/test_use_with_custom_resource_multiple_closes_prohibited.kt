// vybe-test: kotlin/kotlin_closeable_use/test_use_with_custom_resource_multiple_closes_prohibited
// origin: languages/kotlin/tests/kotlin/test_kotlin_closeable_use.rs

import java.io.Closeable

        class Counted : Closeable {
            var closeCount = 0
            override fun close() { closeCount++ }
        }

        fun main() {
            val tracked = Counted()
            tracked.use {
                println(tracked.closeCount)
            }
            println(tracked.closeCount)
            try {
                tracked.close()
                println("extra")
            } catch (e: Exception) {
                println("err")
            }
            println(tracked.closeCount)
        }

