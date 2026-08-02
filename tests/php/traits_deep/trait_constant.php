<?php
// vybe-test: php/traits_deep/trait_constant
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait Configurable {
    const DEFAULT_TIMEOUT = 30;
    const MAX_RETRIES = 3;
}
class HttpClient { use Configurable; }
echo HttpClient::DEFAULT_TIMEOUT . ',' . HttpClient::MAX_RETRIES;
