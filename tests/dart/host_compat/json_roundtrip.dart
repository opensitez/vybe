// vybe-test: dart/host_compat/json_roundtrip
// origin: languages/dart/tests/dart/test_host_compat.rs

var data = json.decode('{"name": "Alice"}');
var str = json.encode(data);

void main() {}
