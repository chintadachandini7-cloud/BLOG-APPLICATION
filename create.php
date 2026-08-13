<?php
include 'db.php';

if (isset($_POST['submit'])) {
    $title = trim($_POST['title']);
    $author = trim($_POST['author']);
    $content = trim($_POST['content']);

    if (!empty($title) && !empty($author) && !empty($content)) {
        $stmt = $conn->prepare("INSERT INTO posts (title, author, content) VALUES (?, ?, ?)");
        $stmt->bind_param("sss", $title, $author, $content);
        $stmt->execute();
        $stmt->close();
        header("Location: index.php");
        exit();
    }
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Create Blog Post</title>
    <style>
        body { font-family: Arial, sans-serif; background: #f4f4f9; padding: 20px; display: flex; justify-content: center; }
        .container { width: 500px; background: #fff; padding: 20px 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h2 { text-align: center; color: #333; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input, textarea { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
        textarea { height: 120px; resize: vertical; }
        button { background: #28a745; color: white; padding: 10px 15px; border: none; border-radius: 4px; cursor: pointer; width: 100%; font-size: 16px; }
        .back-link { display: block; margin-top: 15px; text-align: center; text-decoration: none; color: #666; }
    </style>
</head>
<body>
    <div class="container">
        <h2>Write a New Post</h2>
        <form method="POST" action="create.php">
            <div class="form-group">
                <label>Title</label>
                <input type="text" name="title" required>
            </div>
            <div class="form-group">
                <label>Author Name</label>
                <input type="text" name="author" required>
            </div>
            <div class="form-group">
                <label>Content</label>
                <textarea name="content" required></textarea>
            </div>
            <button type="submit" name="submit">Publish Post</button>
        </form>
        <a href="index.php" class="back-link">← Back to Home</a>
    </div>
</body>
</html>