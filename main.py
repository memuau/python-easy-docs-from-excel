from easy_docs_from_excel.TemplateFiller import TemplateFiller

tf = TemplateFiller.from_yaml('config_to_fill.yml') # Remember to change the name of the file if needed
tf.make_all_docs()
