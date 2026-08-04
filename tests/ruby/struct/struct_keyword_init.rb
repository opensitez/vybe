# vybe-test: ruby/struct/struct_keyword_init
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Config = Struct.new(:host, :port, keyword_init: true)
c = Config.new(host: "localhost", port: 8080)
puts c.host
puts c.port
