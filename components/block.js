polarity.export = PolarityComponent.extend({
  details: Ember.computed.alias('block.data.details'),
  timezone: Ember.computed('Intl', function () {
    return Intl.DateTimeFormat().resolvedOptions().timeZone;
  }),
  activeTab: 'documents',
  init: function () {
    if (!this.get('results')) {
      if (this.get('details.documents.length') > 0) {
        this.set('results', this.get('details.documents'));
        this.set('activeTab', 'documents');
      } else {
        this.set('results', this.get('details.pages'));
        this.set('activeTab', 'pages');
      }
    }
    this._super(...arguments);
  },
  actions: {
    changeTab: function (tab) {
      this.set('activeTab', tab);
      if (tab === 'pages') {
        this.set('results', this.get('details.pages'));
      } else {
        this.set('results', this.get('details.documents'));
      }
    }
  }
});
