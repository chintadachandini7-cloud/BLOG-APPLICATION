<?php
include 'db.php';
$result = $conn->query("SELECT * FROM posts ORDER BY id DESC");
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Simple Blogging Platform</title>
    <style>
        body { font-family: Arial, sans-serif; background: #f4f4f9; padding: 20px; display: flex; justify-content: center; }
        .container { width: 600px; background: #fff; padding: 20px 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h1 { text-align: center; color: #333; }
        .top-btn { display: inline-block; background: #007bff; color: white; padding: 10px 15px; text-decoration: none; border-radius: 4px; margin-bottom: 20px; }
        .post { border-bottom: 1px solid #ddd; padding-bottom: 15px; margin-bottom: 15px; }
        .post h2 { margin: 0 0 5px 0; color: #007bff; }
        .meta { font-size: 12px; color: #777; margin-bottom: 10px; }
        .post p { color: #555; line-height: 1.5; }
    </style>
</head>
<body>
    <div class="container">
        <h1>My Simple Blog</h1>
        <a href="create.php" class="top-btn">+ Create New Post</a>

        <?php if ($result && $result->num_rows > 0): ?>
            <?php while ($row = $result->fetch_assoc()): ?>
                <div class="post">
                    <h2><?php echo htmlspecialchars($row['title']); ?></h2>
                    <div class="meta">By <strong><?php echo htmlspecialchars($row['author']); ?></strong> on <?php echo $row['created_at']; ?></div>
                    <p><?php echo nl2br(htmlspecialchars($row['content'])); ?></p>
                </div>
            <?php endwhile; ?>
        <?php else: ?>
            <p>No blog posts found. Create your first post!</p>
        <?php endif; ?>
    </div>
</body>
</html>