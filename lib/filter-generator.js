/**
 * Thunderbird Message Filters Generator
 * Converts parsed Outlook rules and resolved folder targets into:
 * 1. Standard msgFilterRules.dat file text
 * 2. XPCOM-compatible filter definition objects
 */

class FilterGenerator {
  /**
   * Escape a string value for use inside a Thunderbird condition term: (attrib,op,value)
   * @param {string} str
   * @returns {string}
   */
  static escapeConditionValue(str) {
    if (!str) return '""';
    const escaped = str.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    // Always wrap in quotes if contains whitespace, commas, or parentheses
    if (/[\s,()"]/g.test(escaped)) {
      return `"${escaped}"`;
    }
    return escaped;
  }

  /**
   * Build condition string for a single rule
   * e.g. "AND (from,contains,alice@example.com) OR (subject,contains,invoice)"
   * @param {object} rule
   * @returns {string}
   */
  static buildConditionString(rule) {
    if (!rule.conditions || rule.conditions.length === 0) {
      return 'ALL';
    }

    const terms = [];

    for (const cond of rule.conditions) {
      switch (cond.type) {
        case 'from':
          if (cond.values && cond.values.length > 0) {
            for (const val of cond.values) {
              terms.push({ op: 'OR', term: `(from,contains,${FilterGenerator.escapeConditionValue(val)})` });
            }
          }
          break;

        case 'to':
          if (cond.values && cond.values.length > 0) {
            for (const val of cond.values) {
              terms.push({ op: 'OR', term: `(to,contains,${FilterGenerator.escapeConditionValue(val)})` });
            }
          }
          break;

        case 'subject':
          if (cond.values && cond.values.length > 0) {
            for (const val of cond.values) {
              terms.push({ op: 'OR', term: `(subject,contains,${FilterGenerator.escapeConditionValue(val)})` });
            }
          }
          break;

        case 'body':
          if (cond.values && cond.values.length > 0) {
            for (const val of cond.values) {
              terms.push({ op: 'OR', term: `(body,contains,${FilterGenerator.escapeConditionValue(val)})` });
            }
          }
          break;

        case 'subject_or_body':
          if (cond.values && cond.values.length > 0) {
            for (const val of cond.values) {
              terms.push({ op: 'OR', term: `(subject,contains,${FilterGenerator.escapeConditionValue(val)})` });
              terms.push({ op: 'OR', term: `(body,contains,${FilterGenerator.escapeConditionValue(val)})` });
            }
          }
          break;

        case 'has_attachment':
          terms.push({ op: 'AND', term: '(status,is,hasAttachment)' });
          break;

        case 'priority':
          terms.push({ op: 'AND', term: `(priority,is,${FilterGenerator.escapeConditionValue(cond.value || 'High')})` });
          break;

        case 'name_in_to':
          terms.push({ op: 'AND', term: '(to,contains,"")' });
          break;

        case 'name_in_cc':
          terms.push({ op: 'AND', term: '(cc,contains,"")' });
          break;

        case 'name_in_to_or_cc':
          terms.push({ op: 'AND', term: '(to_or_cc,contains,"")' });
          break;

        default:
          break;
      }
    }

    if (terms.length === 0) {
      return 'ALL';
    }

    // Combine terms into Thunderbird condition format
    // The first term is prefixed with AND or OR, subsequent terms with their logical connector
    return terms.map((t, idx) => {
      const prefix = (idx === 0) ? 'AND' : t.op;
      return `${prefix} ${t.term}`;
    }).join(' ');
  }

  /**
   * Convert rules and resolved folder URIs into msgFilterRules.dat file content.
   *
   * @param {Array<object>} rules Normalized rules
   * @param {Map<string, { folder: object, uri: string }>} folderMap Map of folderName -> { uri }
   * @param {object} [options]
   * @param {boolean} [options.logging=true]
   * @param {string} [options.defaultFolderUri=''] Fallback folder URI
   * @returns {string} msgFilterRules.dat content
   */
  static generateMsgFilterRulesDat(rules, folderMap = new Map(), options = {}) {
    const logging = options.logging !== false ? 'yes' : 'no';
    const lines = [
      'version="9"',
      `logging="${logging}"`
    ];

    for (const rule of rules) {
      lines.push(`name="${(rule.name || '未命名規則').replace(/"/g, '\\"')}"`);
      lines.push(`enabled="${rule.enabled ? 'yes' : 'no'}"`);
      lines.push('type="17"'); // 1 (InboxRule) + 16 (Manual) = 17

      // Actions
      if (rule.actions && rule.actions.length > 0) {
        for (const act of rule.actions) {
          switch (act.type) {
            case 'move_to_folder': {
              const folderInfo = folderMap.get(act.folderName);
              const uri = (folderInfo && folderInfo.uri) || options.defaultFolderUri || `mailbox://nobody@Local%20Folders/${encodeURIComponent(act.folderName)}`;
              lines.push('action="Move to folder"');
              lines.push(`actionValue="${uri}"`);
              break;
            }

            case 'copy_to_folder': {
              const folderInfo = folderMap.get(act.folderName);
              const uri = (folderInfo && folderInfo.uri) || options.defaultFolderUri || `mailbox://nobody@Local%20Folders/${encodeURIComponent(act.folderName)}`;
              lines.push('action="Copy to folder"');
              lines.push(`actionValue="${uri}"`);
              break;
            }

            case 'delete':
              lines.push('action="Delete"');
              break;

            case 'mark_read':
              lines.push('action="Mark read"');
              break;

            case 'mark_flagged':
              lines.push('action="Mark flagged"');
              break;

            case 'stop_execution':
              lines.push('action="Stop execution"');
              break;

            case 'forward':
              if (act.recipients && act.recipients.length > 0) {
                lines.push('action="Forward"');
                lines.push(`actionValue="${act.recipients[0]}"`);
              }
              break;

            case 'change_priority':
              lines.push('action="Change priority"');
              lines.push(`actionValue="${act.priority || 'High'}"`);
              break;

            default:
              break;
          }
        }
      } else {
        // Default action if none specified
        lines.push('action="Stop execution"');
      }

      // Condition
      const conditionStr = FilterGenerator.buildConditionString(rule);
      lines.push(`condition="${conditionStr}"`);
    }

    return lines.join('\n') + '\n';
  }

  /**
   * Convert normalized rules to structured objects suitable for WebExtension Experiment API
   * @param {Array<object>} rules
   * @param {Map<string, { folder: object, uri: string }>} folderMap
   * @returns {Array<object>}
   */
  static toExperimentPayload(rules, folderMap = new Map()) {
    return rules.map(rule => {
      const condition = FilterGenerator.buildConditionString(rule);
      const actions = [];

      for (const act of rule.actions) {
        if (act.type === 'move_to_folder') {
          const fInfo = folderMap.get(act.folderName);
          actions.push({
            type: 1, // MoveToFolder
            actionStr: 'Move to folder',
            targetFolderUri: fInfo ? fInfo.uri : ''
          });
        } else if (act.type === 'copy_to_folder') {
          const fInfo = folderMap.get(act.folderName);
          actions.push({
            type: 16, // CopyToFolder
            actionStr: 'Copy to folder',
            targetFolderUri: fInfo ? fInfo.uri : ''
          });
        } else if (act.type === 'delete') {
          actions.push({ type: 3, actionStr: 'Delete' });
        } else if (act.type === 'mark_read') {
          actions.push({ type: 4, actionStr: 'Mark read' });
        } else if (act.type === 'mark_flagged') {
          actions.push({ type: 7, actionStr: 'Mark flagged' });
        } else if (act.type === 'stop_execution') {
          actions.push({ type: 11, actionStr: 'Stop execution' });
        } else if (act.type === 'forward' && act.recipients && act.recipients.length > 0) {
          actions.push({ type: 10, actionStr: 'Forward', strValue: act.recipients[0] });
        } else if (act.type === 'change_priority') {
          actions.push({ type: 2, actionStr: 'Change priority', priority: act.priority || 'High' });
        }
      }

      return {
        name: rule.name,
        enabled: rule.enabled,
        filterType: 17, // InboxRule (1) | Manual (16)
        condition,
        actions
      };
    });
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { FilterGenerator };
}
