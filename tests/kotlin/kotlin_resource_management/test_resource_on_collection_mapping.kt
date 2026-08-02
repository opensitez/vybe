// vybe-test: kotlin/kotlin_resource_management/test_resource_on_collection_mapping
// origin: languages/kotlin/tests/kotlin/test_kotlin_resource_management.rs

class LogClose : AutoCloseable {
            var tag = ""
            override fun close() { tag = "closed" }
        }

        fun main() {
            val logs = listOf("a", "b")
            val out = logs.joinToString(",") { value ->
                val resource = LogClose()
                resource.use {
                    println(value)
                    value
                }
            }
            println(out)
        }

