use crate::message::Message;
use crate::state::EditorState;
use iced::widget::{button, column, container, row, text, mouse_area};
use iced::{Element, Length, Color};
use iced::widget::canvas::{self, Cache, Frame, Geometry, Path, Stroke};
use iced::{Point, Rectangle, Renderer, Size, Theme};
use irys_forms::{Control, ControlType};

pub struct Designer {
    cache: Cache,
}

impl Designer {
    pub fn new() -> Self {
        Self {
            cache: Cache::new(),
        }
    }

    pub fn view<'a>(&'a self, state: &'a EditorState) -> Element<'a, Message> {
        let canvas = canvas(DesignerCanvas { state })
            .width(Length::Fill)
            .height(Length::Fill);

        let canvas_with_mouse = mouse_area(canvas)
            .on_press(|pos| {
                Message::CanvasClicked(pos.x as i32, pos.y as i32)
            });

        let form_name = state.current_form.as_ref()
            .map(|s| s.as_str())
            .unwrap_or("No Form");

        let designer_view = column![
            row![
                button(text("📝 View Code").size(13))
                    .on_press(Message::ViewCode)
                    .padding(8),
                text("Form Designer").size(16),
                text(format!("({})", form_name))
                    .size(13)
                    .style(|_theme| {
                        text::Style {
                            color: Some(Color::from_rgb(0.5, 0.5, 0.5)),
                        }
                    }),
            ]
            .spacing(10)
            .padding(10),
            container(canvas_with_mouse)
                .width(Length::Fill)
                .height(Length::Fill)
                .style(|_theme| {
                    container::Style {
                        background: Some(iced::Background::Color(Color::WHITE)),
                        border: iced::Border {
                            color: Color::from_rgb(0.8, 0.8, 0.8),
                            width: 1.0,
                            ..Default::default()
                        },
                        ..Default::default()
                    }
                }),
        ]
        .spacing(5);

        container(designer_view)
            .width(Length::Fill)
            .height(Length::Fill)
            .into()
    }

    pub fn clear_cache(&mut self) {
        self.cache.clear();
    }
}

impl Default for Designer {
    fn default() -> Self {
        Self::new()
    }
}

struct DesignerCanvas<'a> {
    state: &'a EditorState,
}

impl<'a> canvas::Program<Message> for DesignerCanvas<'a> {
    type State = ();

    fn draw(
        &self,
        _state: &Self::State,
        renderer: &Renderer,
        _theme: &Theme,
        bounds: Rectangle,
        _cursor: iced::mouse::Cursor,
    ) -> Vec<Geometry> {
        let mut frame = Frame::new(renderer, bounds.size());

        // Draw form background
        frame.fill_rectangle(
            Point::new(0.0, 0.0),
            Size::new(bounds.width, bounds.height),
            Color::from_rgb(0.95, 0.95, 0.95),
        );

        // Draw grid lines
        let grid_size = 20.0;
        let grid_color = Color::from_rgb(0.85, 0.85, 0.85);

        let mut x = 0.0;
        while x < bounds.width {
            let line = Path::line(Point::new(x, 0.0), Point::new(x, bounds.height));
            frame.stroke(&line, Stroke::default().with_width(0.5).with_color(grid_color));
            x += grid_size;
        }

        let mut y = 0.0;
        while y < bounds.height {
            let line = Path::line(Point::new(0.0, y), Point::new(bounds.width, y));
            frame.stroke(&line, Stroke::default().with_width(0.5).with_color(grid_color));
            y += grid_size;
        }

        // Draw all controls from the current form
        if let Some(form) = self.state.get_current_form() {
            for control in &form.controls {
                let is_selected = self.state.selected_control == Some(control.id);
                draw_control(&mut frame, control, is_selected);
            }
        }

        vec![frame.into_geometry()]
    }
}

fn draw_control(frame: &mut Frame, control: &Control, is_selected: bool) {
    let bounds = control.bounds;
    let x = bounds.x as f32;
    let y = bounds.y as f32;
    let width = bounds.width as f32;
    let height = bounds.height as f32;

    // Choose colors based on control type
    let (bg_color, text_color) = match control.control_type {
        ControlType::Button => {
            if control.is_enabled() {
                (Color::from_rgb(0.88, 0.88, 0.88), Color::BLACK)
            } else {
                (Color::from_rgb(0.7, 0.7, 0.7), Color::from_rgb(0.5, 0.5, 0.5))
            }
        }
        ControlType::TextBox => (Color::WHITE, Color::BLACK),
        ControlType::Label => (Color::from_rgb(0.95, 0.95, 0.95), Color::BLACK),
        ControlType::CheckBox | ControlType::RadioButton => {
            (Color::from_rgb(0.95, 0.95, 0.95), Color::BLACK)
        }
        _ => (Color::from_rgb(0.9, 0.9, 0.9), Color::BLACK),
    };

    // Draw control background
    frame.fill_rectangle(
        Point::new(x, y),
        Size::new(width, height),
        bg_color,
    );

    // Draw border
    let border_color = if is_selected {
        Color::from_rgb(0.0, 0.4, 1.0)
    } else {
        Color::from_rgb(0.0, 0.0, 0.0)
    };

    let border = Path::rectangle(Point::new(x, y), Size::new(width, height));
    frame.stroke(
        &border,
        Stroke::default()
            .with_width(if is_selected { 2.5 } else { 1.0 })
            .with_color(border_color),
    );

    // Draw selection handles if selected
    if is_selected {
        let handle_size = 6.0;
        let handles = [
            (x - handle_size/2.0, y - handle_size/2.0), // Top-left
            (x + width/2.0 - handle_size/2.0, y - handle_size/2.0), // Top
            (x + width - handle_size/2.0, y - handle_size/2.0), // Top-right
            (x - handle_size/2.0, y + height/2.0 - handle_size/2.0), // Left
            (x + width - handle_size/2.0, y + height/2.0 - handle_size/2.0), // Right
            (x - handle_size/2.0, y + height - handle_size/2.0), // Bottom-left
            (x + width/2.0 - handle_size/2.0, y + height - handle_size/2.0), // Bottom
            (x + width - handle_size/2.0, y + height - handle_size/2.0), // Bottom-right
        ];

        for (hx, hy) in handles {
            frame.fill_rectangle(
                Point::new(hx, hy),
                Size::new(handle_size, handle_size),
                Color::WHITE,
            );
            let handle_border = Path::rectangle(
                Point::new(hx, hy),
                Size::new(handle_size, handle_size)
            );
            frame.stroke(&handle_border, Stroke::default().with_width(1.0).with_color(Color::BLACK));
        }
    }

    // Draw checkbox/radio button indicator
    match control.control_type {
        ControlType::CheckBox => {
            let box_size = 14.0;
            let box_x = x + 5.0;
            let box_y = y + height / 2.0 - box_size / 2.0;

            frame.fill_rectangle(
                Point::new(box_x, box_y),
                Size::new(box_size, box_size),
                Color::WHITE,
            );
            let checkbox_border = Path::rectangle(Point::new(box_x, box_y), Size::new(box_size, box_size));
            frame.stroke(&checkbox_border, Stroke::default().with_width(1.0).with_color(Color::BLACK));
        }
        ControlType::RadioButton => {
            let circle_size = 14.0;
            let circle_x = x + 5.0 + circle_size / 2.0;
            let circle_y = y + height / 2.0;

            let circle = Path::circle(Point::new(circle_x, circle_y), circle_size / 2.0);
            frame.fill(&circle, Color::WHITE);
            frame.stroke(&circle, Stroke::default().with_width(1.0).with_color(Color::BLACK));
        }
        _ => {}
    }

    // Draw control text/caption
    let display_text = control
        .get_caption()
        .or_else(|| control.get_text())
        .unwrap_or(&control.name);

    let text_x = match control.control_type {
        ControlType::CheckBox | ControlType::RadioButton => x + 25.0,
        ControlType::Button => x + width / 2.0 - (display_text.len() as f32 * 3.5),
        _ => x + 5.0,
    };

    let text_y = y + height / 2.0 - 7.0;

    frame.fill_text(canvas::Text {
        content: display_text.to_string(),
        position: Point::new(text_x, text_y),
        color: text_color,
        size: 14.0.into(),
        ..Default::default()
    });

    // Draw control type label in corner for debugging
    if is_selected {
        let type_label = format!("[{}]", control.control_type.as_str());
        frame.fill_text(canvas::Text {
            content: type_label,
            position: Point::new(x + 2.0, y + 2.0),
            color: Color::from_rgb(0.5, 0.5, 0.5),
            size: 10.0.into(),
            ..Default::default()
        });
    }
}
