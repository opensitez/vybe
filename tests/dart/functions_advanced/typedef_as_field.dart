// vybe-test: dart/functions_advanced/typedef_as_field
// origin: languages/dart/tests/dart/test_functions_advanced.rs

typedef Handler = void Function(String); class Server { Handler? onMessage; }

void main() {}
