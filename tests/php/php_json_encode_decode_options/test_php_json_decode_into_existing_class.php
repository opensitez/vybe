<?php
// vybe-test: php/php_json_encode_decode_options/test_php_json_decode_into_existing_class
// origin: languages/php/tests/php/test_php_json_encode_decode_options.rs
// vybe-test-mode: compile

class UserDto {
    public string $name;
    public int $age;
}

$json = '{"name":"Alice","age":30}';
$dto = json_decode($json);
echo "$dto->name is $dto->age";
