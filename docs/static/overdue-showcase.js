/* Visual presentation only. Demo paths never call business or messaging APIs. */
(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const money = n => Number(n || 0).toLocaleString('zh-CN', {minimumFractionDigits:2, maximumFractionDigits:2});
  const rules = [
    {level:'', action:'暂不动作', tone:'—', tier:'none'},
    {level:'低', action:'持续观察', tone:'首次提醒', tier:'low'},
    {level:'中', action:'电话提示经办人', tone:'首次提醒', tier:'medium'},
    {level:'高', action:'生成催款函', tone:'正式催收', tier:'high'},
    {level:'高', action:'追加停止供货警示', tone:'严正警示', tier:'critical'}
  ];
  const band = days => days >= 30 ? 4 : days >= 15 ? 3 : days >= 5 ? 2 : days >= 1 ? 1 : 0;
  window.odResolveRisk = st => {
    const periods = (st.periods || []).filter(p => Number(p.unpaid) > 0 && (p.is_overdue || Number(p.overdue_days) > 0));
    const p = periods.sort((a,b) => Number(b.overdue_days) - Number(a.overdue_days))[0];
    const decisions = (st.risk_agent || {}).decisions || [];
    const rank = {高:3, 中:2, 低:1};
    const d = decisions.filter(x => rank[x.risk_level]).slice().sort((a,b) => rank[b.risk_level] - rank[a.risk_level])[0];
    if (d) return d;
    if (!p) return null;
    const r = rules[band(Number(p.overdue_days))];
    return {risk_level:r.level, action:r.action, letter_tone:r.tone, period_key:p.period_key,
      reason:`已逾期 ${Number(p.overdue_days)} 天，按合同阈值${r.action}。`};
  };
  let latestRisk = {days:0, rule:rules[0]};
  window.odRenderSpotlight = (st, top) => {
    const sum = st.summary || {};
    const amount = Number(sum.current_overdue || 0);
    const periods = (st.periods || []).filter(p => Number(p.unpaid) > 0 && (p.is_overdue || Number(p.overdue_days) > 0));
    const days = Math.max(0, ...periods.map(p => Number(p.overdue_days) || 0));
    const idx = amount > 0 ? band(days) : 0;
    const rule = rules[idx];
    const level = top && top.risk_level || rule.level;
    const tier = amount <= 0 ? 'none' : days >= 30 ? 'critical' : level === '高' ? 'high' : level === '中' ? 'medium' : 'low';
    latestRisk = {days, rule, tier};
    $('odRiskBanner').removeAttribute('style');
    $('odRiskTitle').removeAttribute('style');
    $('odRiskIcon').removeAttribute('style');
    $('odRiskBanner').dataset.tier = tier;
    $('odRiskIcon').textContent = amount > 0 ? '!' : '✓';
    $('odRiskTitle').textContent = amount > 0 ? `¥ ${money(amount)}` : '当前无逾期';
    $('odRiskKicker').textContent = amount > 0 ? '已识别逾期余额 · 规则核算' : '风险监测 · 持续守候';
    $('odRiskPill').textContent = amount > 0 ? `${level || '待判定'}风险${days >= 30 ? ' · 停供警示' : ''}` : '持续监控';
    $('odRiskDays').textContent = amount > 0 ? String(days || '—') : '0';
    $('odRiskDesc').textContent = amount > 0 ? (top && top.reason || `按合同阈值${rule.action}。`) : '未到起息日或已完成回款，系统持续监控。';
    $('odRiskMeta').textContent = `建议动作：${top && top.action || rule.action} · ${Number(sum.overdue_period_count || periods.length)} 个逾期周期 · 数据基准：${sum.last_scan || sum.checked_at || '当前状态'}`;
    document.querySelectorAll('#odRiskScale [data-rule]').forEach(el => {
      const current = Number(el.dataset.rule) === idx;
      el.classList.toggle('is-current', current);
      if (current) el.setAttribute('aria-current','step'); else el.removeAttribute('aria-current');
    });
    if (sum.virtual_demo) $('odFlowMode').textContent = '数据快照 · 下方可预览自动流转';
    else $('odFlowMode').textContent = `每 ${sum.monitor_interval_minutes || 10} 分钟复查 · 动态路径示意`;
  };
  function flowState(root, active, done = []) {
    root.querySelectorAll('.od-node').forEach((el,i) => {
      el.classList.toggle('is-active', i === active);
      el.classList.toggle('is-done', done.includes(i));
      if (i === active) el.setAttribute('aria-current', 'step'); else el.removeAttribute('aria-current');
    });
  }
  let previewTimer = null;
  $('odFlowPreview').addEventListener('click', () => {
    const button = $('odFlowPreview');
    const flow = $('odLiveFlow');
    const path = latestRisk.days >= 15 ? [0,1,2,3,4] : latestRisk.days >= 5 ? [0,1,2,4] : [0,1,2];
    let cursor = 0;
    button.disabled = true;
    function tick() {
      const active = path[cursor];
      flowState(flow, active, path.slice(0,cursor));
      const captions = ['定时任务唤醒 · 预览中','读取合同、账期与回款依据','按逾期天数命中风险档位','生成催款函 · 流程示意','推送经办人 · 消息模拟'];
      $('odFlowCaption').textContent = captions[active];
      previewTimer = setTimeout(() => {
        cursor++;
        if (cursor < path.length) tick();
        else {
          flowState(flow, -1, path);
          $('odFlowCaption').textContent = latestRisk.days < 5 ? '持续观察 · 当前档位不生成催款函，不推送催办消息' : latestRisk.days < 15 ? '电话提示经办人 · 当前档位不生成催款函（预览结束）' : '催款函送达经办人，等待本人确认（预览结束）';
          button.disabled = false;
          previewTimer = null;
        }
      }, 950);
    }
    tick();
  });

  // A coherent example: one cycle, one principal balance, actual contract thresholds.
  const amount = '135,906.45';
  const file = '催款函-A项目-20261022.docx';
  const stages = [
    {label:'定时扫描',date:'2026-10-08',days:0,level:'none',node:0,duration:4800,title:'定时复查，把风险找出来',
      note:'每 10 分钟复查结算周期、合同条件与回款记录。先核对付款条件，再识别逾期余额。', action:'账期 + 合同 + 回款证据 → 规则核算', evidence:'sources'},
    {label:'首次逾期',date:'2026-10-08',days:1,level:'low',node:2,duration:5200,title:'到达起息日，捕捉首次逾期',
      note:'该周期未收 135,906.45 元，起息日当天计为逾期第 1 天。未到期的其他周期继续监控。',action:'低风险 · 持续观察，暂不生成催款函',evidence:'low'},
    {label:'电话提示',date:'2026-10-12',days:5,level:'medium',node:2,duration:5200,title:'跨过 5 天阈值，提醒经办人',
      note:'仍未收到回款，风险进入中档。系统建议先电话提示经办人，推动主动跟进。',action:'中风险 · 电话提示经办人 / 首次提醒',evidence:'medium'},
    {label:'风险升级',date:'2026-10-22',days:15,level:'high',node:2,duration:5800,title:'逾期满 15 天，正式催收',
      note:'风险由中升高。相同账期、相同余额，按合同阈值进入催款函生成节点。',action:'高风险 · 生成催款函 / 正式催收',evidence:'high'},
    {label:'函件生成',date:'2026-10-22',days:15,level:'high',node:3,duration:6200,title:'从核算依据，到一份催款函',
      note:'将该周期结算、回款与未收余额带入函件，交由经办人复核后转发甲方。',action:'函件待复核 · 每个金额都有依据',evidence:'letter'},
    {label:'飞书送达',date:'2026-10-22',days:15,level:'high',node:4,duration:6200,title:'催办消息，送到责任人手中',
      note:'飞书卡片携带项目、账期、逾期余额和催款函附件。展示模拟送达与待确认状态。',action:'模拟推送 · 李经理 / 待本人确认',evidence:'feishu'},
    {label:'本人确认',date:'2026-10-22',days:15,level:'high',node:4,duration:5600,title:'经办人确认，让责任闭环',
      note:'确认表示本次催办已收到、可继续处理；不代表欠款结清。系统继续监测后续回款。',action:'催办已确认 · 逾期余额仍持续监控',evidence:'confirmed'},
    {label:'停供警示',date:'2026-11-06',days:30,level:'critical',node:2,duration:6500,title:'若满 30 天仍未回款，继续升级',
      note:'追加停止供货警示，使用严正警示话术。是否实际停供，由有权限的负责人复核决定。',action:'高风险 · 追加停止供货警示 / 人工决策',evidence:'critical'}
  ];
  const dialog = $('odDemoOverlay');
  let index = 0, timer = null, paused = false, ended = false, started = 0, remaining = 0, focusBefore;
  const speed = () => Number($('odDemoSpeed').value);
  const clear = () => { clearTimeout(timer); timer = null; };
  const source = (title, body, meta) => `<div class="od-source"><b>${title}</b>${body}<small>${meta}</small></div>`;
  function evidence(stage) {
    const body = $('odEvidenceBody');
    $('odEvidenceTitle').textContent = stage.evidence === 'letter' ? '催款函 · 预览摘录' : ['feishu','confirmed'].includes(stage.evidence) ? '飞书 · 消息卡片' : '风险判断 · 依据可追溯';
    $('odEvidenceBadge').textContent = ['feishu','confirmed'].includes(stage.evidence) ? '模拟消息' : '示例数据';
    if (stage.evidence === 'letter') {
      body.innerHTML = `<article class="od-letter"><h4>催款函</h4><p>致：示例项目方有限公司</p><p>截至 2026年10月22日，周期 2026年7月21日—8月20日结算金额 304,637.68 元，已付 168,731.23 元，未收余额 <mark>${amount} 元</mark>。</p><p>该周期于 2026年10月8日起息，现已逾期 15 天。请积极筹措资金，尽快履行合同约定的付款义务。</p><p>本摘录列示该周期未收余额；资金占用费以核算结果另行复核。</p><div class="od-letter-sign">示例供应链管理有限公司<br>2026年10月22日</div></article>`;
    } else if (['feishu','confirmed'].includes(stage.evidence)) {
      const confirmed = stage.evidence === 'confirmed';
      body.innerHTML = `<div class="od-feishu-card"><div class="od-feishu-header"><b>飞书 · 风险提醒助手</b><span>模拟推送</span></div><div class="od-feishu-content"><h4>A项目 · 逾期催办</h4><p>经办人：李经理<br>周期：2026/07/21—08/20<br>逾期余额：<b>¥ ${amount}</b><br>逾期天数：15 天 · 高风险</p><div class="od-feishu-file">▤ ${file}</div><div class="od-message-receipt">${confirmed ? '本人已确认 · 催办处理留痕' : '已送达（模拟）· 待本人确认'}</div></div></div>`;
    } else if (stage.evidence === 'sources') {
      body.innerHTML = source('01 / 合同规则','起息日：2026-10-08','规则引擎确定付款条件，不提前计息') + source('02 / 结算周期','结算金额 ¥ 304,637.68','周期 2026/07/21—08/20') + source('03 / 回款证据','已付 ¥ 168,731.23','差额核算 → 未收 ¥ 135,906.45');
    } else {
      const r = rules[band(stage.days)];
      body.innerHTML = source('01 / 事实核算',`逾期 ${stage.days} 天 · 未收 ¥ ${amount}`,'起息日 2026-10-08 · 起息日当天为第 1 天') + source('02 / 命中阈值',`${stage.days >= 30 ? '≥30 天' : stage.days >= 15 ? '15–29 天' : stage.days >= 5 ? '5–14 天' : '1–4 天'} → ${r.level}风险`,`建议动作：${r.action} · 话术：${r.tone}`) + source('03 / 责任边界',stage.days >= 30 ? '发出停供警示，交负责人决定' : stage.days >= 15 ? '生成催款函，交经办人复核' : '持续监控，按阈值升级','金额与天数由规则核算 · AI 辅助组织措辞 · 人工确认处置');
    }
    body.scrollTop = 0;
  }
  function controls() {
    dialog.classList.toggle('od-is-paused', paused || ended);
    $('odDemoPauseBtn').textContent = ended ? '重新播放' : paused ? '继续播放' : '暂停播放';
    $('odDemoPrevBtn').disabled = index === 0;
    $('odDemoNextBtn').disabled = index === stages.length - 1;
    $('odDemoStatus').textContent = ended ? '本轮演示结束 · 催办闭环后，风险仍持续监控' : `${paused ? '已暂停 · 可逐步讲解' : '自动播放中'} · ${index + 1} / ${stages.length}`;
  }
  function render(i) {
    index = i;
    const s = stages[i];
    const idx = band(s.days);
    dialog.style.setProperty('--scene-duration', `${s.duration * speed()}ms`);
    $('odShowScene').dataset.level = s.level;
    $('odSceneDay').textContent = s.date;
    $('odSceneRisk').textContent = i === 0 ? '准备扫描' : `${rules[idx].level}风险${idx === 4 ? ' · 停供警示' : ''}`;
    $('odSceneTitle').textContent = s.title;
    $('odSceneNote').textContent = s.note;
    $('odSceneAmount').innerHTML = `${i === 0 ? '—' : amount}<em> 元</em>`;
    $('odSceneDays').textContent = i === 0 ? '—' : String(s.days);
    $('odSceneAction').textContent = s.action;
    $('odSceneLadder').querySelectorAll('span').forEach((el,n) => el.classList.toggle('active', n === idx));
    const story = $('odSceneStory');
    story.classList.remove('od-reveal');
    void story.offsetWidth;
    story.classList.add('od-reveal');
    flowState($('odDemoFlow'), s.node, Array.from({length:s.node},(_,n)=>n));
    evidence(s);
    $('odDemoChapters').innerHTML = stages.map((x,n) => `<button class="od-chapter ${n === i ? 'active' : n < i ? 'done' : ''}" type="button" data-step="${n}" aria-label="第 ${n+1} 步：${x.label}" ${n === i ? 'aria-current="step"' : ''}><b>${String(n+1).padStart(2,'0')}</b>${x.label}</button>`).join('');
    $('odDemoSlider').value = String(i);
    $('odDemoSlider').setAttribute('aria-valuetext', `${i + 1} / ${stages.length} ${s.label}`);
    controls();
  }
  function schedule(ms = stages[index].duration * speed()) {
    clear();
    if (paused || ended || !dialog.open || document.hidden) return;
    remaining = ms;
    started = performance.now();
    timer = setTimeout(() => {
      if (index === stages.length - 1) { ended = true; controls(); return; }
      render(index + 1); schedule();
    }, ms);
  }
  function seek(i) {
    clear(); paused = true; ended = false;
    render(Math.max(0,Math.min(stages.length - 1,i)));
    remaining = stages[index].duration * speed();
  }
  window.openOverdueDemo = () => {
    if (dialog.open) return;
    focusBefore = document.activeElement;
    dialog.showModal();
    dialog.scrollTop = 0;
    paused = false; ended = false;
    render(0); schedule();
    $('odDemoPauseBtn').focus({preventScroll:true});
  };
  window.closeOdDemo = ev => {
    if (ev) { ev.preventDefault(); ev.stopPropagation(); }
    clear();
    dialog.close();
  };
  window.odDemoTogglePause = () => {
    if (ended) { window.odDemoRestart(); return; }
    if (!paused) {
      remaining = Math.max(0, remaining - (performance.now() - started));
      clear(); paused = true;
    } else { paused = false; schedule(remaining); }
    controls();
  };
  window.odDemoRestart = () => { paused = false; ended = false; render(0); schedule(); };
  window.odDemoNext = () => seek(index + 1);
  window.odDemoSeek = value => seek(Number(value));
  $('odDemoPrevBtn').addEventListener('click', () => seek(index - 1));
  $('odDemoChapters').addEventListener('click', ev => { const btn = ev.target.closest('[data-step]'); if (btn) seek(Number(btn.dataset.step)); });
  $('odDemoSpeed').addEventListener('change', () => { render(index); remaining = stages[index].duration * speed(); schedule(); });
  dialog.addEventListener('close', () => { clear(); if (focusBefore && focusBefore.isConnected) focusBefore.focus({preventScroll:true}); });
  dialog.addEventListener('click', ev => { if (ev.target === dialog) { const b = dialog.getBoundingClientRect(); if (ev.clientX < b.left || ev.clientX > b.right || ev.clientY < b.top || ev.clientY > b.bottom) window.closeOdDemo(); } });
  document.addEventListener('visibilitychange', () => {
    if (!dialog.open || paused || ended) return;
    if (document.hidden) { remaining = Math.max(0, remaining - (performance.now() - started)); clear(); }
    else schedule(remaining);
  });
})();
