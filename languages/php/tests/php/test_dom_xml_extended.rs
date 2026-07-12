//! DOM and SimpleXML APIs not covered in `test_dom_xml.rs` `php_cases!` block.

crate::php_cases! {
    dom_document_create_text_node => {
        r#"<?php
$doc = new DOMDocument();
$text = $doc->createTextNode('hello');
echo $text->textContent;
"#,
        ["hello"]
    };

    dom_element_set_attribute => {
        r#"<?php
$doc = new DOMDocument();
$el = $doc->createElement('item');
$el->setAttribute('id', '42');
echo $el->getAttribute('id');
"#,
        ["42"]
    };

    dom_element_remove_attribute => {
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<a x="1"/>');
$doc->documentElement->removeAttribute('x');
echo $doc->documentElement->hasAttribute('x') ? 'yes' : 'no';
"#,
        ["no"]
    };

    dom_get_element_by_id => {
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<root><item id="main">x</item></root>');
$doc->documentElement->setIdAttribute('id', true);
echo $doc->getElementById('main')->textContent;
"#,
        ["x"]
    };

    dom_node_child_nodes_length => {
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<root><a/><b/></root>');
echo $doc->documentElement->childNodes->length;
"#,
        ["2"]
    };

    dom_node_first_child_tag => {
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<list><first/><second/></list>');
echo $doc->documentElement->firstChild->nodeName;
"#,
        ["first"]
    };

    dom_node_next_sibling => {
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<p><a/><b/></p>');
$a = $doc->getElementsByTagName('a')->item(0);
echo $a->nextSibling->nodeName;
"#,
        ["b"]
    };

    dom_cdata_section => {
        r#"<?php
$doc = new DOMDocument();
$cdata = $doc->createCDATASection('raw<data>');
echo $cdata->textContent;
"#,
        ["raw<data>"]
    };

    dom_comment_node => {
        r#"<?php
$doc = new DOMDocument();
$comment = $doc->createComment('note');
echo $comment->textContent;
"#,
        ["note"]
    };

    dom_save_xml_without_xml_declaration => {
        r#"<?php
$doc = new DOMDocument('1.0', 'UTF-8');
$doc->formatOutput = false;
$root = $doc->createElement('r');
$doc->appendChild($root);
$xml = $doc->saveXML($root);
echo str_starts_with(trim($xml), '<r') ? 'ok' : 'no';
"#,
        ["ok"]
    };

    dom_xpath_query_count => {
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<items><i/><i/><i/></items>');
$xp = new DOMXPath($doc);
echo $xp->query('//i')->length;
"#,
        ["3"]
    };

    dom_xpath_text_extract => {
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<p><t>hi</t></p>');
$xp = new DOMXPath($doc);
echo $xp->query('//t')->item(0)->textContent;
"#,
        ["hi"]
    };

    simplexml_load_file_from_string_wrapper => {
        r#"<?php
$xml = simplexml_load_string('<?xml version="1.0"?><root val="1"/>');
echo (string)$xml['val'];
"#,
        ["1"]
    };

    simplexml_add_child => {
        r#"<?php
$xml = simplexml_load_string('<root/>');
$xml->addChild('child', 'text');
echo (string)$xml->child;
"#,
        ["text"]
    };

    simplexml_as_xml_roundtrip => {
        r#"<?php
$xml = simplexml_load_string('<a><b/></a>');
echo str_contains($xml->asXML(), '<b') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    simplexml_count_children => {
        r#"<?php
$xml = simplexml_load_string('<r><c/><c/></r>');
echo count($xml->c);
"#,
        ["2"]
    };

    simplexml_xpath_query => {
        r#"<?php
$xml = simplexml_load_string('<book><title>A</title></book>');
$hits = $xml->xpath('//title');
echo (string)$hits[0];
"#,
        ["A"]
    };

    libxml_clear_errors_after_bad_parse => {
        r#"<?php
libxml_use_internal_errors(true);
simplexml_load_string('<bad');
libxml_clear_errors();
echo count(libxml_get_errors());
"#,
        ["0"]
    };

    dom_import_simplexml_preserves_text => {
        r#"<?php
$sx = simplexml_load_string('<n>9</n>');
$node = dom_import_simplexml($sx);
echo $node->textContent;
"#,
        ["9"]
    };

    dom_document_encoding_property => {
        r#"<?php
$doc = new DOMDocument('1.0', 'UTF-8');
echo $doc->encoding;
"#,
        ["UTF-8"]
    };

    dom_element_tag_name_upper => {
        r#"<?php
$doc = new DOMDocument();
$el = $doc->createElement('item');
echo $el->tagName;
"#,
        ["item"]
    };

    dom_node_parent_node => {
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<root><leaf/></root>');
$leaf = $doc->getElementsByTagName('leaf')->item(0);
echo $leaf->parentNode->nodeName;
"#,
        ["root"]
    };

    dom_append_document_fragment => {
        r#"<?php
$doc = new DOMDocument();
$frag = $doc->createDocumentFragment();
$frag->appendXML('<x/>');
$root = $doc->createElement('root');
$root->appendChild($frag);
$doc->appendChild($root);
echo $doc->getElementsByTagName('x')->length;
"#,
        ["1"]
    };

    simplexml_getname_root => {
        r#"<?php
$xml = simplexml_load_string('<catalog/>');
echo $xml->getName();
"#,
        ["catalog"]
    };

    dom_load_html_fragment => {
        r#"<?php
$doc = new DOMDocument();
@$doc->loadHTML('<p>html</p>');
echo $doc->getElementsByTagName('p')->item(0)->textContent;
"#,
        ["html"]
    };
}
