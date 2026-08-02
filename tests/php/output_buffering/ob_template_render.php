<?php
// vybe-test: php/output_buffering/ob_template_render
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

function renderTemplate(string $title, array $items): string {
    ob_start(); ?>
<h1><?= htmlspecialchars($title) ?></h1>
<ul>
<?php foreach ($items as $item): ?>
  <li><?= htmlspecialchars($item) ?></li>
<?php endforeach; ?>
</ul>
<?php return ob_get_clean();
}
$html = renderTemplate('My List', ['Apple', 'Banana', 'Cherry']);
echo strlen($html) > 0 ? 'rendered' : 'empty';
