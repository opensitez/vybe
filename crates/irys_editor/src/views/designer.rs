use crate::message::Message;
use crate::state::EditorState;
use iced::widget::{button, column, container, row, text, Canvas};
use iced::{Color, Element, Length, mouse, Point, Rectangle, Renderer, Size, Theme};
use iced::widget::canvas::{self, Cache, Frame, Geometry, Path, Stroke};
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
        let canvas = Canvas::new(DesignerCanvas { state, cache: &self.cache })
            .width(Length::Fill)
            .height(Length::Fill);

        let designer_view = column![
            row![
                button(text("View Code")).on_press(Message::ViewCode),
                text("Form Designer").size(16),
            ]
            .spacing(10),
            canvas,
        ]
        .spacing(5)
        .padding(5);

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
    cache: &'a Cache,
}

pub struct CanvasState {
    pub last_click_time: Option<std::time::Instant>,
    pub last_click_pos: Option<(i32, i32)>,
    pub last_clicked_control: Option<uuid::Uuid>,
}

impl Default for CanvasState {
    fn default() -> Self {
        Self {
            last_click_time: None,
            last_click_pos: None,
            last_clicked_control: None,
        }
    }
}

impl<'a> canvas::Program<Message> for DesignerCanvas<'a> {
    type State = CanvasState;

    fn update(
        &self,
        state: &mut Self::State,
        event: canvas::Event,
        bounds: Rectangle,
        cursor: mouse::Cursor,
    ) -> (canvas::event::Status, Option<Message>) {
        match event {
            canvas::Event::Mouse(mouse_event) => {
                match mouse_event {
                    mouse::Event::ButtonPressed(mouse::Button::Left) => {
                        if let Some(position) = cursor.position_in(bounds) {
                            let x = position.x as i32;
                            let y = position.y as i32;

                            // Check for double-click
                            let now = std::time::Instant::now();
                            let is_double_click = if let Some(last_time) = state.last_click_time {
                                let elapsed = now.duration_since(last_time);
                                if elapsed.as_millis() < 500 {
                                    // Within 500ms
                                    if let Some((last_x, last_y)) = state.last_click_pos {
                                        // Within 5 pixels
                                        (x - last_x).abs() < 5 && (y - last_y).abs() < 5
                                    } else {
                                        false
                                    }
                                } else {
                                    false
                                }
                            } else {
                                false
                            };

                            // Find control at position
                            let clicked_control = if let Some(form) = self.state.get_current_form() {
                                form.find_control_at(x, y).map(|c| c.id)
                            } else {
                                None
                            };

                            let message = if is_double_click && clicked_control.is_some() && clicked_control == state.last_clicked_control {
                                // Double-click on control
                                Message::ControlDoubleClicked(clicked_control.unwrap())
                            } else {
                                // Single click
                                Message::CanvasClicked(x, y)
                            };

                            // Update state for next click
                            state.last_click_time = Some(now);
                            state.last_click_pos = Some((x, y));
                            state.last_clicked_control = clicked_control;

                            return (
                                canvas::event::Status::Captured,
                                Some(message)
                            );
                        }
                    }
                    mouse::Event::CursorMoved { .. } => {
                        if let Some(position) = cursor.position_in(bounds) {
                            // Check if we're dragging
                            if self.state.drag_state.is_some() {
                                return (
                                    canvas::event::Status::Captured,
                                    Some(Message::DesignerMouseMove(position.x as i32, position.y as i32))
                                );
                            }
                        }
                    }
                    mouse::Event::ButtonReleased(mouse::Button::Left) => {
                        if self.state.drag_state.is_some() {
                            return (
                                canvas::event::Status::Captured,
                                Some(Message::DesignerMouseUp)
                            );
                        }
                    }
                    _ => {}
                }
            }
            _ => {}
        }

        (canvas::event::Status::Ignored, None)
    }

    fn draw(
        &self,
        _state: &Self::State,
        renderer: &Renderer,
        _theme: &Theme,
        bounds: Rectangle,
        _cursor: mouse::Cursor,
    ) -> Vec<Geometry> {
        let geometry = self.cache.draw(renderer, bounds.size(), |frame| {
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
                    draw_control(frame, control, is_selected);
                }
            }
        });

        vec![geometry]
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
        ControlType::RichTextBox => (Color::WHITE, Color::BLACK),
        ControlType::Label => (Color::from_rgb(0.95, 0.95, 0.95), Color::BLACK),
        ControlType::CheckBox | ControlType::RadioButton => {
            (Color::from_rgb(0.95, 0.95, 0.95), Color::BLACK)
        }
        ControlType::WebBrowser => (Color::from_rgb(0.98, 0.98, 1.0), Color::BLACK),
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
}
