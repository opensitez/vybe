//! Code editor engine — wraps vybe_widgets::TextEditor with LSP diagnostic support.

use lsp_types::Diagnostic;
use std::collections::HashSet;
use vybe_widgets::language::LanguageDef;
use vybe_widgets::text_editor::TextEditor;
pub use vybe_widgets::text_editor::{LexerState, TokenSpan};

/// Editor wraps TextEditor and adds LSP-specific diagnostics.
pub struct Editor {
    pub inner: TextEditor,
    #[allow(dead_code)]
    pub diagnostics: Vec<Diagnostic>,
}

#[allow(dead_code)]
impl Editor {
    pub fn from_text(text: &str, lang: &LanguageDef) -> Self {
        Self {
            inner: TextEditor::from_text(text, lang),
            diagnostics: Vec::new(),
        }
    }

    // Delegate all methods to inner
    pub fn rope(&self) -> &ropey::Rope {
        self.inner.rope()
    }
    pub fn save_snapshot(&mut self, line: usize, col: usize) {
        self.inner.save_snapshot(line, col);
    }
    pub fn undo(&mut self, cur_line: usize, cur_col: usize) -> Option<(String, usize, usize)> {
        self.inner.undo(cur_line, cur_col)
    }
    pub fn redo(&mut self, cur_line: usize, cur_col: usize) -> Option<(String, usize, usize)> {
        self.inner.redo(cur_line, cur_col)
    }
    pub fn delete_range(&mut self, start_byte: usize, end_byte: usize, lang: &LanguageDef) {
        self.inner.delete_range(start_byte, end_byte, lang);
    }
    pub fn slice(&self, start: usize, end: usize) -> String {
        self.inner.slice(start, end)
    }
    pub fn insert_str(&mut self, byte_pos: usize, s: &str, lang: &LanguageDef) {
        self.inner.insert_str(byte_pos, s, lang);
    }
    pub fn insert_newline(&mut self, byte_pos: usize, lang: &LanguageDef) {
        self.inner.insert_newline(byte_pos, lang);
    }
    pub fn replace_all(&mut self, from: &str, to: &str, lang: &LanguageDef) {
        self.inner.replace_all(from, to, lang);
    }
    pub fn retokenize_all(&mut self, lang: &LanguageDef) {
        self.inner.retokenize_all(lang);
    }
    pub fn retokenize_range(&mut self, start_line: usize, end_line: usize, lang: &LanguageDef) {
        self.inner.retokenize_range(start_line, end_line, lang);
    }
    pub fn tokenize_all(&self) -> Vec<TokenSpan> {
        self.inner.tokenize_all()
    }
    pub fn toggle_fold(&mut self, line_idx: usize) {
        self.inner.toggle_fold(line_idx);
    }
    pub fn find_matching_bracket(
        &self,
        line_idx: usize,
        col: usize,
        lang: &LanguageDef,
    ) -> Option<(usize, usize)> {
        self.inner.find_matching_bracket(line_idx, col, lang)
    }
    pub fn folds(&self) -> &Vec<(usize, usize)> {
        self.inner.folds()
    }
    pub fn move_line_up(&mut self, line_idx: usize) {
        self.inner.move_line_up(line_idx);
    }
    pub fn move_line_down(&mut self, line_idx: usize) {
        self.inner.move_line_down(line_idx);
    }
    pub fn duplicate_line(&mut self, line_idx: usize) {
        self.inner.duplicate_line(line_idx);
    }
    pub fn insert_string(
        &mut self,
        byte_pos: usize,
        text: &str,
        lang: &LanguageDef,
    ) -> (usize, usize) {
        self.inner.insert_string(byte_pos, text, lang)
    }
}

// Expose inner fields via Deref-like accessors for renderer.rs compatibility
#[allow(dead_code)]
impl Editor {
    pub fn line_tokens(&self) -> &Vec<Vec<TokenSpan>> {
        &self.inner.line_tokens
    }
    pub fn line_states(&self) -> &Vec<LexerState> {
        &self.inner.line_states
    }
    pub fn collapsed_starts(&self) -> &HashSet<usize> {
        &self.inner.collapsed_starts
    }
}

// Keep the old field-access patterns working via a Deref-like approach
impl std::ops::Deref for Editor {
    type Target = TextEditor;
    fn deref(&self) -> &TextEditor {
        &self.inner
    }
}
impl std::ops::DerefMut for Editor {
    fn deref_mut(&mut self) -> &mut TextEditor {
        &mut self.inner
    }
}
