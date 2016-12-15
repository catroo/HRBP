<?php
echo preg_match('/^√([\x4e00-\x9fa5]{2})\s+.*$/U', '√已育  未育', $matches1);
echo preg_match('/^√(.+)$/', '已育  √未育   ', $matches2);
print_r($matches1);
print_r($matches2);
