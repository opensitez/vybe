// vybe-test: kotlin/kotlin_path_and_files/test_path_iterate_name_count
// origin: languages/kotlin/tests/kotlin/test_kotlin_path_and_files.rs

import java.nio.file.Paths

        fun main() {
            val path = Paths.get("/x/y/z/file.data")
            println(path.nameCount)
            var parts = ""
            for (part in path) {
                parts += part.fileName.toString() + "/"
            }
            println(parts)
            println(path.getName(1).toString())
        }

