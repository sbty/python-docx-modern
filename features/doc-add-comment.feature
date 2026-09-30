Feature: Add a comment to a document
  In order add a comment to a document
  As a developer using python-docx
  I need a way to add a comment specifying both its content and its reference


  Scenario: Document.add_comment(runs, text, author, initials)
    Given a document having a comments part
     When I assign comment = document.add_comment(runs, "A comment", "John Doe", "JD")
     Then comment is a Comment object
      And comment.text == "A comment"
      And comment.author == "John Doe"
      And comment.initials == "JD"


  Scenario: Document.add_comment() defines the comment styles it refers to
    Given a default document having a paragraph with a run
     When I assign comment = document.add_comment(runs, "A comment", "John Doe", "JD")
     Then the document defines the "CommentReference" character style
      And the document defines the "CommentText" paragraph style


  Scenario: Comments.add_comment() defines the comment styles it refers to
    Given a default document having a paragraph with a run
     When I assign comment = document.comments.add_comment()
     Then the document defines the "CommentReference" character style
      And the document defines the "CommentText" paragraph style
