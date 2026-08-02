// vybe-test: kotlin/kotlin_lazy_delegates/test_delegates_not_null_requires_assignment_before_read
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_delegates.rs

import kotlin.properties.Delegates

        fun main() {
            class Holder {
                var name: String by Delegates.notNull()
            }

            val holder = Holder()
            try {
                holder.name.length
                println("ready")
            } catch (e: IllegalStateException) {
                println(e::class.simpleName)
            }
            holder.name = "ok"
            println(holder.name)
        }

