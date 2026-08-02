<?php
// vybe-test: php/php_spl_observer_subject_pattern/test_php_spl_observer_subject_event_notification
// origin: languages/php/tests/php/test_php_spl_observer_subject_pattern.rs

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

class NewsPublisher implements SplSubject {
    private SplObjectStorage $observers;
    public string $latestArticle = "";

    public function __construct() {
        $this->observers = new SplObjectStorage();
    }

    public function attach(SplObserver $observer): void {
        $this->observers->attach($observer);
    }

    public function detach(SplObserver $observer): void {
        $this->observers->detach($observer);
    }

    public function notify(): void {
        foreach ($this->observers as $observer) {
            $observer->update($this);
        }
    }

    public function publish(string $article): void {
        $this->latestArticle = $article;
        $this->notify();
    }
}

class Reader implements SplObserver {
    public string $received = "";
    public function update(SplSubject $subject): void {
        if ($subject instanceof NewsPublisher) {
            $this->received = $subject->latestArticle;
        }
    }
}

$publisher = new NewsPublisher();
$reader = new Reader();
$publisher->attach($reader);

$publisher->publish("PHP 8.4 Released!");
echo "Reader received: {$reader->received}";

__vybe_check(ob_get_clean(), "Reader received: PHP 8.4 Released!");
