<?php
// vybe-test: php/declare/strict_types_named_args
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
function makeTag(string $tag, string $content, bool $self_close = false): string {
    if ($self_close) return "<$tag />";
    return "<$tag>$content</$tag>";
}
echo makeTag(content: 'Hello', tag: 'p');
