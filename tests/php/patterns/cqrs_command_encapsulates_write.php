<?php
// vybe-test: php/patterns/cqrs_command_encapsulates_write
// origin: languages/php/tests/php/test_patterns.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

interface Command {}
class CreateUserCommand implements Command {
    public function __construct(public readonly string $name, public readonly string $email) {}
}
class CommandBus {
    private $handlers = [];
    public function register(string $commandClass, callable $handler): void {
        $this->handlers[$commandClass] = $handler;
    }
    public function dispatch(Command $cmd): void {
        $class = get_class($cmd);
        if (!isset($this->handlers[$class])) throw new \Exception("no handler for $class");
        ($this->handlers[$class])($cmd);
    }
}
$bus = new CommandBus();
$bus->register(CreateUserCommand::class, function(CreateUserCommand $cmd) {
    echo 'created:' . $cmd->name . ':' . $cmd->email;
});
$bus->dispatch(new CreateUserCommand('Alice', 'alice@example.com'));

__vybe_check(ob_get_clean(), "created:Alice:alice@example.com");
