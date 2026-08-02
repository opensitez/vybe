<?php
// vybe-test: php/reflection/reflection_doc_comment
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

/** @param int $n The number */
function documented(int $n): int { return $n * 2; }
$rf = new ReflectionFunction('documented');
$doc = $rf->getDocComment();
echo $doc !== false ? 'has doc' : 'no doc';
