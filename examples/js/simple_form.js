let form = gui.createForm("Test");
gui.setProperty(form, "Width", 300);
gui.setProperty(form, "Height", 200);

gui.addControl(form, "Label", "lbl", 20, 20, 260, 30);
gui.setProperty("lbl", "Text", "Not clicked yet");

gui.addControl(form, "Button", "btn", 20, 60, 100, 30);
gui.setProperty("btn", "Text", "Click");

gui.onEvent("btn", "Click", () => {
    gui.setProperty("lbl", "Text", "Clicked!");
});

gui.runApplication(form);
