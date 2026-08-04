// vybe-test: dart/dart_apis/try_catch
// origin: languages/dart/tests/dart/test_dart_apis.rs

try { throw Exception('oops'); } catch (e) { print(e); } finally { print('done'); }

void main() {}
