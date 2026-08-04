// vybe-test: dart/dart_apis/factory_constructor
// origin: languages/dart/tests/dart/test_dart_apis.rs

class Logger { static Logger? _instance; factory Logger() { _instance ??= Logger(); return _instance; } Logger(); }

void main() {}
