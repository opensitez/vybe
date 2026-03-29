use eframe::egui;
use crate::state::{EditorState, RunStatus, View};
use crate::panels;
use crate::panels::properties::PropertiesTab;

pub struct VybeApp {
    pub state: EditorState,
    pub show_output: bool,
    pub properties_tab: PropertiesTab,
}

impl VybeApp {
    pub fn new(_cc: &eframe::CreationContext, cli_path: Option<std::path::PathBuf>) -> Self {
        Self {
            state: EditorState::new(cli_path),
            show_output: false,
            properties_tab: PropertiesTab::Properties,
        }
    }
}

impl eframe::App for VybeApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        self.state.poll_run();

        // Show output panel when running or done
        if matches!(self.state.run_status, RunStatus::Running | RunStatus::Done(_)) {
            self.show_output = true;
        }

        // Request repaint while running so we poll
        if matches!(self.state.run_status, RunStatus::Running) {
            ctx.request_repaint_after(std::time::Duration::from_millis(200));
        }

        // ── Menu bar ──────────────────────────────────────────────────
        egui::TopBottomPanel::top("menu").show(ctx, |ui| {
            egui::menu::bar(ui, |ui| {
                ui.menu_button("File", |ui| {
                    if ui.button("New Project").clicked() { self.state.new_project(); ui.close_menu(); }
                    if ui.button("Open Project…").clicked() { self.state.open_project_dialog(); ui.close_menu(); }
                    ui.separator();
                    if ui.button("Save").clicked() { self.state.save_project(); ui.close_menu(); }
                    if ui.button("Save As…").clicked() { self.state.save_project_as(); ui.close_menu(); }
                    ui.separator();
                    if ui.button("Exit").clicked() { ctx.send_viewport_cmd(egui::ViewportCommand::Close); }
                });
                ui.menu_button("Edit", |ui| {
                    if ui.button("Undo  Ctrl+Z").clicked() { self.state.undo(); ui.close_menu(); }
                    if ui.button("Redo  Ctrl+Y").clicked() { self.state.redo(); ui.close_menu(); }
                    ui.separator();
                    if ui.button("Delete  Del").clicked() { self.state.delete_selected(); ui.close_menu(); }
                });
                ui.menu_button("View", |ui| {
                    ui.checkbox(&mut self.state.show_project_explorer, "Project Explorer");
                    ui.checkbox(&mut self.state.show_toolbox, "Toolbox");
                    ui.checkbox(&mut self.state.show_properties, "Properties");
                    ui.checkbox(&mut self.show_output, "Output");
                });
                ui.menu_button("Run", |ui| {
                    let running = matches!(self.state.run_status, RunStatus::Running);
                    if !running && ui.button("▶  Start  F5").clicked() {
                        self.state.run();
                        ui.close_menu();
                    }
                    if running && ui.button("■  Stop").clicked() {
                        self.state.stop();
                        ui.close_menu();
                    }
                });
            });
        });

        // ── Toolbar ───────────────────────────────────────────────────
        egui::TopBottomPanel::top("toolbar").show(ctx, |ui| {
            ui.horizontal(|ui| {
                let running = matches!(self.state.run_status, RunStatus::Running);

                // Run / Stop
                if running {
                    if ui.button("■ Stop").clicked() { self.state.stop(); }
                } else {
                    if ui.button("▶ Run").clicked() { self.state.run(); self.show_output = true; }
                }

                ui.separator();

                // View toggles
                let in_designer = self.state.view == View::FormDesigner && self.state.current_form.is_some();
                let in_code = self.state.view == View::CodeEditor;

                if ui.selectable_label(in_designer, "🗒 Designer").clicked() {
                    if self.state.current_form.is_some() {
                        self.state.view = View::FormDesigner;
                    }
                }
                if ui.selectable_label(in_code, "📝 Code").clicked() {
                    self.state.view = View::CodeEditor;
                }

                ui.separator();

                // Save shortcut
                if ui.button("💾 Save").clicked() { self.state.save_project(); }

                // Project name
                if let Some(proj) = &self.state.project {
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        ui.label(egui::RichText::new(&proj.name).weak());
                    });
                }
            });
        });

        // ── Output panel (bottom) ─────────────────────────────────────
        if self.show_output {
            egui::TopBottomPanel::bottom("output")
                .resizable(true)
                .min_height(80.0)
                .default_height(150.0)
                .show(ctx, |ui| {
                    ui.horizontal(|ui| {
                        ui.heading("Output");
                        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                            if ui.small_button("✕").clicked() { self.show_output = false; }
                        });
                    });
                    ui.separator();
                    panels::output::show(ui, &mut self.state);
                });
        }

        // ── Left: Project Explorer ────────────────────────────────────
        if self.state.show_project_explorer {
            egui::SidePanel::left("project_explorer")
                .resizable(true)
                .min_width(150.0)
                .default_width(200.0)
                .show(ctx, |ui| {
                    egui::ScrollArea::vertical().show(ui, |ui| {
                        panels::project_explorer::show(ui, &mut self.state);
                    });
                });
        }

        // ── Left: Toolbox (only in form designer) ────────────────────
        let in_designer = self.state.view == View::FormDesigner && self.state.current_form.is_some();
        if in_designer && self.state.show_toolbox {
            egui::SidePanel::left("toolbox")
                .resizable(true)
                .min_width(130.0)
                .default_width(155.0)
                .show(ctx, |ui| {
                    egui::ScrollArea::vertical().show(ui, |ui| {
                        panels::toolbox::show(ui, &mut self.state);
                    });
                });
        }

        // ── Right: Properties (only in form designer) ─────────────────
        if in_designer && self.state.show_properties {
            egui::SidePanel::right("properties")
                .resizable(true)
                .min_width(180.0)
                .default_width(220.0)
                .show(ctx, |ui| {
                    egui::ScrollArea::vertical().show(ui, |ui| {
                        panels::properties::show(ui, &mut self.state, &mut self.properties_tab);
                    });
                });
        }

        // ── Central area ──────────────────────────────────────────────
        egui::CentralPanel::default().show(ctx, |ui| {
            match self.state.view {
                View::FormDesigner => panels::form_designer::show(ui, &mut self.state),
                View::CodeEditor => panels::code_editor::show(ui, &mut self.state),
            }
        });

        // Global keyboard shortcuts
        ctx.input(|i| {
            if i.key_pressed(egui::Key::F5) && !matches!(self.state.run_status, RunStatus::Running) {
                self.state.run();
                self.show_output = true;
            }
        });
    }
}
