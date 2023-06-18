from easy_docs_from_excel.TemplateFiller import TemplateFiller

tf = TemplateFiller.from_yaml('myconfig.yml')
tf.make_all_docs()