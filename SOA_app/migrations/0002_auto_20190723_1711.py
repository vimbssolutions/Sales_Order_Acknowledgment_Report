# Generated by Django 2.2 on 2019-07-23 11:41

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('SOA_app', '0001_initial'),
    ]

    operations = [
        migrations.AlterField(
            model_name='sale_order_summery_v',
            name='Shipping_Method',
            field=models.CharField(max_length=64, null=True),
        ),
    ]
