<?php
// vybe-test: php/enums_deep/enum_static_method_factory
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Environment: string {
    case Development = 'dev';
    case Staging     = 'staging';
    case Production  = 'prod';
    public static function fromEnvVar(): self {
        return self::tryFrom(getenv('APP_ENV') ?: '') ?? self::Development;
    }
    public function isProduction(): bool { return $this === self::Production; }
}
$env = Environment::fromEnvVar();
echo $env->name;
echo $env->isProduction() ? ':prod' : ':not prod';
