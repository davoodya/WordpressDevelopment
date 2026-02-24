### Usage Guide
1. Copy product-exporter.php to you host in wordpress root directory (./public_html, ./root/, ...)
2. Run product-exporter.php with redirection
```php 
php product-exporter.php > products.csv
```
3. Download Products.csv from your host 

### Note:
This script export below fields:
1. name
2. special-sale-price

from:
- only one variants from all variants of a product
- and do this for all products

> so export name and special-sale-price from first variant of all products