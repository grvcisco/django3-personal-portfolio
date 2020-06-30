
from django.urls import path
from . import views

app_name ="blog"
urlpatterns = [
    path('', views.all_blogs, name='all_blogs'),
    path('<int:blog_id>/', views.detail, name="detail"),
    path(r'^xls/$', views.export_xls, name='export_xls'),

]
