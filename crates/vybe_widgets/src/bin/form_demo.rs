//! Demo form application — shows the toolkit used like VB/WinForms.
//!
//! Run with:  cargo run -p vybe_widgets --bin form_demo

use cosmic_text::{FontSystem, SwashCache};
use tiny_skia::Pixmap;
use winit::window::CursorIcon;

use vybe_widgets::{
    Application, Button, Checkbox, Form, Label, PanelWidget, TextInput, WidgetEvent,
    layout::{KeyEvent, LayoutRect, MouseEvent, RenderContext},
    run_app,
};

struct FormDemoApp {
    form: Form,
    font_system: FontSystem,
    swash_cache: SwashCache,
    scale: f32,
    width: f32,
    height: f32,
}

impl FormDemoApp {
    fn new() -> Self {
        Self {
            form: Form::new("Demo Form"),
            font_system: FontSystem::new(),
            swash_cache: SwashCache::new(),
            scale: 2.0,
            width: 0.0,
            height: 0.0,
        }
    }

    fn build_form(&mut self) {
        self.form = Form::new("Demo Form");
        self.form
            .set_rect(LayoutRect::new(0.0, 0.0, self.width, self.height));

        // Title label
        let mut title = Label::new("Vybe Toolkit Demo");
        title.font_size = 18.0;
        title.colors.foreground = (30, 30, 30, 255);
        self.form.add_control(title, 20.0, 15.0, 300.0, 30.0);

        // Name label + text input
        self.form
            .add_control(Label::new("Name:"), 20.0, 60.0, 60.0, 24.0);
        self.form.add_control(
            TextInput::new()
                .with_name("name")
                .with_placeholder("Enter your name"),
            90.0,
            60.0,
            250.0,
            26.0,
        );

        // Email label + text input
        self.form
            .add_control(Label::new("Email:"), 20.0, 100.0, 60.0, 24.0);
        self.form.add_control(
            TextInput::new()
                .with_name("email")
                .with_placeholder("you@example.com"),
            90.0,
            100.0,
            250.0,
            26.0,
        );

        // Password label + text input
        self.form
            .add_control(Label::new("Password:"), 20.0, 140.0, 80.0, 24.0);
        self.form.add_control(
            TextInput::new()
                .with_name("password")
                .with_placeholder("••••••••")
                .with_password(),
            110.0,
            140.0,
            230.0,
            26.0,
        );

        // Checkboxes
        self.form.add_control(
            Checkbox::new("Remember me").with_name("remember"),
            20.0,
            185.0,
            160.0,
            22.0,
        );
        self.form.add_control(
            Checkbox::new("Accept terms").with_name("terms"),
            20.0,
            215.0,
            160.0,
            22.0,
        );

        // Buttons
        let mut submit = Button::new("Submit");
        submit.name = "submit".to_string();
        submit.colors.background = (0, 102, 204, 255);
        submit.colors.foreground = (255, 255, 255, 255);
        submit.colors.border = (0, 80, 170, 255);
        self.form.add_control(submit, 20.0, 260.0, 100.0, 32.0);

        let mut cancel = Button::new("Cancel");
        cancel.name = "cancel".to_string();
        self.form.add_control(cancel, 130.0, 260.0, 100.0, 32.0);

        // Status label
        let mut status = Label::new("Ready");
        status.colors.foreground = (100, 100, 100, 255);
        self.form.add_control(status, 20.0, 310.0, 320.0, 20.0);
    }

    fn handle_events(&mut self) {
        let events = self.form.drain_events();
        for ev in events {
            match ev {
                WidgetEvent::ButtonClicked(name) => {
                    println!("[event] Button clicked: {}", name);
                }
                WidgetEvent::CheckboxToggled(name, checked) => {
                    println!("[event] Checkbox '{}' = {}", name, checked);
                }
                WidgetEvent::TextChanged(name, value) => {
                    println!("[event] Text '{}' = '{}'", name, value);
                }
                _ => {}
            }
        }
    }
}

impl Application for FormDemoApp {
    fn on_init(&mut self, width: f32, height: f32, scale: f32) {
        self.width = width;
        self.height = height;
        self.scale = scale;
        self.build_form();
    }

    fn on_resize(&mut self, width: f32, height: f32) {
        self.width = width;
        self.height = height;
        self.form.set_rect(LayoutRect::new(0.0, 0.0, width, height));
    }

    fn render(&mut self, pixmap: &mut Pixmap, scale: f32) {
        self.handle_events();

        let mut ctx = RenderContext {
            pixmap,
            font_system: &mut self.font_system,
            swash_cache: &mut self.swash_cache,
            scale,
        };
        self.form.render(&mut ctx);
    }

    fn handle_mouse(&mut self, event: MouseEvent) -> bool {
        self.form.handle_mouse(&event)
    }

    fn handle_key(&mut self, event: KeyEvent) -> bool {
        self.form.handle_key(&event)
    }

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        self.form.handle_scroll(delta, x, y)
    }

    fn cursor_icon(&self) -> CursorIcon {
        self.form.cursor_at(0.0, 0.0)
    }
}

fn main() {
    run_app("Form Demo", 400, 360, 2.0, FormDemoApp::new());
}
