<?php
// vybe-test: php/traits_deep/trait_abstract_template_method
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait Report {
    abstract protected function gatherData(): array;
    abstract protected function formatRow(array $row): string;
    public function generate(): string {
        $rows = array_map([$this, 'formatRow'], $this->gatherData());
        return implode("\n", $rows);
    }
}
class SalesReport {
    use Report;
    protected function gatherData(): array { return [['item' => 'Widget', 'qty' => 5], ['item' => 'Gadget', 'qty' => 3]]; }
    protected function formatRow(array $row): string { return "{$row['item']}: {$row['qty']}"; }
}
echo (new SalesReport())->generate();
