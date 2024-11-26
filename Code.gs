function onOpen() {
  //SlidesApp.getUi().createMenu('Slides Generator')
  //  .addItem('Generate Improve Phase Slides', 'createImprovePhaseSlides')
  //  .addToUi();

  SlidesApp.getUi().removeMenu('Slides Generator')
}

function createImprovePhaseSlides() {
  const slideData = [
    {
      title: "Introduction to Improve Phase",
      content: [
        "Purpose: Highlight methods focused on process improvements.",
        "Key Goals: Waste reduction, cycle time improvement, quality and delivery enhancement."
      ]
    },
    {
      title: "Design of Experiments (DOE)",
      content: [
        "DOE is a structured approach for identifying cause-and-effect relationships in processes.",
        "Key Terms:",
        "- Factors: Controlled variables.",
        "- Levels: Set values for each factor.",
        "- Responses: Measured outcomes."
      ]
    },
    {
      title: "Implementation Planning",
      content: [
        "Stages of Implementation:",
        "1. Proof of Concept",
        "2. Try-Storming (testing ideas quickly)",
        "3. Pilot Testing (readiness check).",
        "Metrics tracking for time and resources."
      ]
    },
    {
      title: "Lean Tools for Waste Elimination",
      content: [
        "Overview of Lean Tools:",
        "- 5S: Organize workspaces.",
        "- Kaizen: Continuous improvement.",
        "- Visual Factory: Using visual indicators."
      ]
    },
    {
      title: "Value Stream Mapping and Kanban",
      content: [
        "Using Value Stream Mapping to identify process inefficiencies and Kanban for demand-driven production control."
      ]
    },
    {
      title: "Poka-Yoke and Standard Work",
      content: [
        "Poka-Yoke: Error-proofing measures to prevent mistakes.",
        "Standard Work: Establishes consistent methods to ensure quality."
      ]
    },
    {
      title: "PDCA Cycle for Continual Improvement",
      content: [
        "PDCA Cycle:",
        "- Plan: Identify improvement area.",
        "- Do: Implement small changes.",
        "- Check: Assess results.",
        "- Act: Standardize successful changes."
      ]
    },
    {
      title: "Conclusion and Q&A",
      content: [
        "Summary:",
        "- Tools and techniques to enhance efficiency and quality.",
        "- Lean principles reduce waste and improve speed.",
        "Q&A: Closing slide for questions."
      ]
    }
  ];

  const presentation = SlidesApp.getActivePresentation();
  
  slideData.forEach(slide => {
    const newSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    
    const titleBox = newSlide.insertTextBox(slide.title);
    titleBox.setTop(50).setLeft(50).setWidth(400).setHeight(50);
    titleBox.getText().getTextStyle().setFontSize(24).setBold(true);
    
    const contentBox = newSlide.insertTextBox(slide.content.join("\n"));
    contentBox.setTop(120).setLeft(50).setWidth(400).setHeight(300);
    contentBox.getText().getTextStyle().setFontSize(16);
  });
  
  SlidesApp.getUi().alert('Improve Phase Slides Created Successfully!');
}
