function LMM_Analysis_v6()
%LMM_ANALYSIS_V6  One-way / two-way ANOVA designs (independent or repeated
%                 measures) analysed with linear (mixed) models, with post hoc
%                 tests and export of all results to Excel.
%
%   Run from the MATLAB command window:   >> LMM_Analysis_v6
%
%   VERSION HISTORY
%     v1  Linear mixed model (fitlme) with post hoc tests and Excel export.
%     v2  Handles designs where some factor combinations have no data
%         (cell-means model, estimable-contrast ANOVA, Design_Check sheet).
%     v3  Asks which design to analyse:
%           1) One-way ANOVA                (independent groups)
%           2) One-way repeated-measures ANOVA
%           3) Two-way ANOVA                (independent groups)
%           4) Two-way repeated-measures ANOVA (both factors within-subject)
%           5) Two-way mixed ANOVA          (one within, one between factor)
%         Lets you name the factors for the output. Checks subject IDs for
%         duplicate measurements and for subjects appearing in more than one
%         group. Day-as-continuous (slopes) option from v2 is not included.
%     v4  Two-way RM ANOVA with a nested layout (e.g. each day belongs to only
%         one training type): random effects become (1|Subject) +
%         (1|Subject:A), so the A effect (e.g. training) is tested against
%         subject-to-subject variation in that effect. The Subject x B term,
%         which cannot be separated from residual error, is omitted.
%     v5  Tukey p-values are now computed exactly from the studentized range
%         distribution (Tukey / Tukey-Kramer, as in Origin, SPSS, GraphPad and
%         R's ptukey) instead of by simulation, so they no longer vary in the
%         third decimal. The simulated max-t adjustment is kept, in its own
%         column, for families that are not all-pairwise (interaction contrasts).
%     v6  Optional comparisons against a reference: after the fit you can pick
%         a reference condition (or several, which are averaged into one
%         baseline) and the conditions to compare with it. Results go to the
%         PH_vs_Reference sheet with Dunnett (many-to-one) adjusted p-values.
%         (Later text-only corrections, no computation changed: the
%         p_SingleStep_maxT description in the output README sheet now states
%         that the fixed seed makes repeat runs identical; "Holm" is now called
%         "Holm-Bonferroni" throughout, so the output columns are named
%         p_Holm_Bonferroni and Significant_Holm_Bonferroni_0p05 instead of
%         p_Holm and Significant_Holm_0p05.)
%
%   MODELS (all factors categorical; effects coding -> Type III tests)
%     One-way ANOVA        : Y ~ A                                   (fitlm)
%     One-way RM ANOVA     : Y ~ A + (1|Subject)                     (fitlme)
%     Two-way ANOVA        : Y ~ A*B                                 (fitlm)
%     Two-way RM ANOVA     : Y ~ A*B + (1|Subject) + (1|Subject:A) + (1|Subject:B)
%       (nested layout, each level of B in only one level of A:
%                          Y ~ A*B + (1|Subject) + (1|Subject:A); and vice versa)
%     Two-way mixed ANOVA  : Y ~ A*B + (1|Subject)       A within, B between
%     With complete, balanced data these give the same F-tests and post hoc
%     results as the classical (sphericity-assumed) ANOVAs. Unlike classical
%     RM ANOVA, subjects with missing values are kept (their other values are
%     used). Mixed models use REML and Satterthwaite degrees of freedom.
%
%   MISSING VALUES
%     Empty cells and text such as NA, N/A, NaN, '.', '-' are treated as
%     missing. Only rows with a missing value in a column used by the analysis
%     are removed; every removed row is listed in the output.
%
%   POST HOC TESTS (t-tests on estimated marginal means, EMMs)
%     One-way : all pairwise level comparisons.
%     Two-way : A main effect, B main effect, A within each level of B,
%               B within each level of A, interaction contrasts, and every
%               A x B cell vs every other cell.
%     P-value adjustments within each family / group: Tukey (exact
%     studentized range; Tukey-Kramer for unequal SEs), Bonferroni, Holm-Bonferroni,
%     Sidak, Hochberg, Benjamini-Hochberg (FDR), Benjamini-Yekutieli,
%     single-step max-t (simulated) and Scheffe.
%
%   REQUIREMENTS: MATLAB R2020b or newer, Statistics and Machine Learning Toolbox.

rng(1);   % reproducible simulated (single-step / Tukey-type) p-values

%% ------------------------------------------------------------------------
%  1. Which analysis?
%  ------------------------------------------------------------------------
typeLabels = { ...
    'One-way ANOVA  (independent groups, one factor)', ...
    'One-way repeated-measures ANOVA  (one within-subject factor)', ...
    'Two-way ANOVA  (independent groups, two factors)', ...
    'Two-way repeated-measures ANOVA  (both factors within-subject)', ...
    'Two-way mixed ANOVA  (one within-subject + one between-subjects factor)'};
kinds = {'oneway', 'oneway_rm', 'twoway', 'twoway_rm', 'mixed'};
[k, ok] = listdlg('PromptString', 'Which analysis do you want to run?', ...
    'ListString', typeLabels, 'SelectionMode', 'single', ...
    'ListSize', [480 110], 'Name', 'Analysis type');
if ~ok, disp('Cancelled.'); return; end
kind      = kinds{k};
typeLabel = typeLabels{k};
twoWay    = ismember(kind, {'twoway', 'twoway_rm', 'mixed'});
isRM      = ismember(kind, {'oneway_rm', 'twoway_rm', 'mixed'});

%% ------------------------------------------------------------------------
%  2. Select and read the data file
%  ------------------------------------------------------------------------
[fn, fp] = uigetfile({'*.xlsx;*.xls;*.xlsm;*.xlsb;*.csv', ...
    'Spreadsheet files (*.xlsx, *.xls, *.xlsm, *.csv)'}, ...
    'Select the Excel file containing your data');
if isequal(fn, 0), disp('No file selected. Exiting.'); return; end
inFile = fullfile(fp, fn);
[~, baseName, ext] = fileparts(inFile);

readArgs = {'VariableNamingRule', 'preserve'};
sheetName = '';
if ~strcmpi(ext, '.csv')
    sh = sheetnames(inFile);
    k = 1;
    if numel(sh) > 1
        [k, ok] = listdlg('PromptString', 'Select the sheet with your data:', ...
            'ListString', cellstr(sh), 'SelectionMode', 'single', ...
            'ListSize', [320 200], 'Name', 'Select sheet');
        if ~ok, disp('Cancelled.'); return; end
    end
    sheetName = char(sh(k));
    readArgs = [readArgs, {'Sheet', sheetName}];
end
try
    raw = readtable(inFile, readArgs{:}, 'TreatAsMissing', missingTokens());
catch
    raw = readtable(inFile, readArgs{:});
end
if height(raw) == 0
    errordlg('The selected file/sheet contains no data rows.', 'LMM'); return;
end
fprintf('Loaded %d rows x %d columns from %s\n', height(raw), width(raw), inFile);

%% ------------------------------------------------------------------------
%  3. Select the columns and name the variables
%  ------------------------------------------------------------------------
vn = raw.Properties.VariableNames;
respVar = pickColumn(vn, 'Select the RESPONSE column (the measured values):');
if isempty(respVar), return; end
bPrompt = '';
switch kind
    case 'oneway'
        aPrompt = 'Select the FACTOR column (group / condition):';
    case 'oneway_rm'
        aPrompt = 'Select the REPEATED FACTOR column (e.g. day, session, condition):';
    case 'twoway'
        aPrompt = 'Select the FIRST FACTOR column:';
        bPrompt = 'Select the SECOND FACTOR column:';
    case 'twoway_rm'
        aPrompt = 'Select the FIRST REPEATED (within-subject) FACTOR column:';
        bPrompt = 'Select the SECOND REPEATED (within-subject) FACTOR column:';
    case 'mixed'
        aPrompt = 'Select the REPEATED (within-subject) FACTOR column, e.g. day:';
        bPrompt = 'Select the BETWEEN-SUBJECTS FACTOR column, e.g. group:';
end
aVar = pickColumn(vn, aPrompt);  if isempty(aVar), return; end
bVar = '';  subjVar = '';
if twoWay
    bVar = pickColumn(vn, bPrompt);  if isempty(bVar), return; end
end
if isRM
    subjVar = pickColumn(vn, 'Select the SUBJECT ID column (e.g. animal ID):');
    if isempty(subjVar), return; end
end
usedCols = {respVar, aVar, bVar, subjVar};
usedCols = usedCols(~cellfun(@isempty, usedCols));
if numel(unique(usedCols)) < numel(usedCols)
    errordlg('Each variable must come from a different column.', 'LMM'); return;
end

prompts = {'Name for the response:', sprintf('Name for the factor in column "%s":', aVar)};
defs    = {respVar, aVar};
if twoWay
    prompts{end+1} = sprintf('Name for the factor in column "%s":', bVar);
    defs{end+1}    = bVar;
end
answ = inputdlg(prompts, 'Names used in the output (optional)', [1 60], defs);
if isempty(answ), answ = defs; end
respName = cleanName(answ{1}, respVar);
nameA    = cleanName(answ{2}, aVar);
nameB    = '';
if twoWay
    nameB = cleanName(answ{3}, bVar);
    if strcmp(nameA, nameB), nameB = [nameB '_2']; end
end

%% ------------------------------------------------------------------------
%  4. Build the analysis table, handle missing values, check the design
%  ------------------------------------------------------------------------
nRaw = height(raw);
T = table();
T.ExcelRow = (1:nRaw)' + 1;                 % assumes headers are in row 1
T.Y  = toNumeric(raw.(respVar));
T.FA = toFactor(raw.(aVar));
if twoWay, T.FB = toFactor(raw.(bVar)); end
if isRM,   T.Subject = toCategorical(raw.(subjVar)); end

missItems = {"Missing/non-numeric " + respName, isnan(T.Y) | ~isfinite(T.Y); ...
             "Missing " + nameA,                isundefined(T.FA)};
if twoWay, missItems(end+1, :) = {"Missing " + nameB, isundefined(T.FB)}; end
if isRM,   missItems(end+1, :) = {"Missing subject ID", isundefined(T.Subject)}; end

reason = strings(nRaw, 1);  drop = false(nRaw, 1);
for r = 1:size(missItems, 1)
    m = missItems{r, 2};
    reason(m) = reason(m) + lower(missItems{r, 1}) + "; ";
    drop = drop | m;
end
missSummary = table(["Rows in file"; string(missItems(:, 1)); "Rows removed (any reason)"; "Rows analysed"], ...
    [nRaw; cellfun(@nnz, missItems(:, 2)); nnz(drop); nnz(~drop)], ...
    'VariableNames', {'Item', 'Count'});
removedRows = table(T.ExcelRow(drop), strip(reason(drop), 'right', ' '), ...
    'VariableNames', {'ExcelRow', 'Reason'});

T = T(~drop, :);
T.FA = removecats(T.FA);
if twoWay, T.FB = removecats(T.FB); end
if isRM,   T.Subject = removecats(T.Subject); end
fprintf('%d rows removed because of missing values; %d rows analysed.\n', nnz(drop), height(T));

aLv = categories(T.FA);  nA = numel(aLv);
if twoWay, bLv = categories(T.FB); nB = numel(bLv); else, bLv = {''}; nB = 1; end
if height(T) < 4 || nA < 2 || (twoWay && nB < 2)
    errordlg('Too few observations or factor levels remain after removing missing values.', 'LMM'); return;
end

checks = cell(0, 2);
notes  = strings(0, 1);

if strcmp(kind, 'mixed')
    % Each subject should belong to exactly one level of the between factor
    [gs, sid] = findgroups(T.Subject);
    nLv = splitapply(@(x) numel(unique(x)), T.FB, gs);
    bad = sid(nLv > 1);
    if ~isempty(bad)
        msg = sprintf(['%d subject ID(s) appear in more than one level of "%s" (e.g. %s).\n\n' ...
            'In a mixed design each subject belongs to one group. If different subjects share an ' ...
            'ID across groups, they can be treated as separate subjects (ID + group).'], ...
            numel(bad), nameB, char(bad(1)));
        a = questdlg(msg, 'Subject IDs', 'Treat as different subjects', 'Cancel', 'Treat as different subjects');
        if ~strcmp(a, 'Treat as different subjects'), disp('Cancelled.'); return; end
        T.Subject = categorical(string(T.Subject) + " [" + string(T.FB) + "]");
        checks(end+1, :) = {'Subjects in more than one group', ...
            sprintf('%d ID(s), treated as different subjects (ID + group): %s', numel(bad), strjoin(string(bad), ', '))};
    else
        checks(end+1, :) = {'Subjects in more than one group', 'None'};
    end
end

if isRM
    % More than one value for the same subject in the same condition?
    if twoWay, gk = findgroups(T.Subject, T.FA, T.FB); else, gk = findgroups(T.Subject, T.FA); end
    cnt  = accumarray(gk, 1);
    dupS = unique(T.Subject(ismember(gk, find(cnt > 1))));
    if ~isempty(dupS)
        msg = sprintf(['%d subject ID(s) have more than one value in the same condition (e.g. %s).\n\n' ...
            'This usually means two subjects share an ID (a typo). If you continue, these values are ' ...
            'treated as repeated measurements of ONE subject. Cancel to fix the IDs in your file first.'], ...
            numel(dupS), char(dupS(1)));
        a = questdlg(msg, 'Duplicate measurements', 'Continue', 'Cancel', 'Cancel');
        if ~strcmp(a, 'Continue'), disp('Cancelled.'); return; end
        checks(end+1, :) = {'Subjects with >1 value in the same condition', ...
            sprintf('%d: %s (treated as the same subject)', numel(dupS), strjoin(string(dupS), ', '))};
    else
        checks(end+1, :) = {'Subjects with >1 value in the same condition', 'None'};
    end
    checks(end+1, :) = {'Number of subjects', sprintf('%d', numel(categories(T.Subject)))};
end

% Which factor combinations contain data?
cellN = zeros(nA, nB);
for i = 1:nA
    for j = 1:nB
        if twoWay
            cellN(i, j) = nnz(T.FA == aLv{i} & T.FB == bLv{j});
        else
            cellN(i, j) = nnz(T.FA == aLv{i});
        end
    end
end
fullDesign = all(cellN(:) > 0);
% Nested layouts: every level of B occurs with only one level of A (B nested
% in A, e.g. each day belongs to one training type), or vice versa.
nestedBinA = twoWay && all(sum(cellN > 0, 1) == 1);
nestedAinB = twoWay && all(sum(cellN > 0, 2) == 1);
[jj, ii] = meshgrid(1:nB, 1:nA);
designCheck = table(string(aLv(ii(:))), cellN(:), 'VariableNames', {nameA, 'N_obs'});
if twoWay
    designCheck = addvars(designCheck, string(bLv(jj(:))), 'After', 1, 'NewVariableNames', nameB);
end
designCheck.HasData = designCheck.N_obs > 0;
if ~fullDesign
    fprintf(['\n%d of %d %s x %s combinations have no data. The full factorial model is not ' ...
        'estimable,\nso the model is fitted on the combinations that exist (cell-means model).\n'], ...
        nnz(cellN == 0), numel(cellN), nameA, nameB);
end

%% ------------------------------------------------------------------------
%  5. Model specification and fit
%  ------------------------------------------------------------------------
reTerm = '';
if isRM
    reTerm = '(1|Subject)';
    if strcmp(kind, 'twoway_rm')
        T.SubjA = categorical(string(T.Subject) + "|" + string(T.FA));
        T.SubjB = categorical(string(T.Subject) + "|" + string(T.FB));
        if nestedBinA && nestedAinB
            reTerm = '(1|Subject)';
            notes(end+1) = "Each level of " + nameA + " pairs with exactly one level of " + nameB + ...
                ", so the two factors carry the same information; only a Subject random intercept was used.";
        elseif nestedBinA
            reTerm = '(1|Subject) + (1|SubjA)';
            notes(end+1) = "Nested layout: each level of " + nameB + " occurs in only one level of " + nameA + ...
                ". Random effects: Subject and Subject x " + nameA + ", so " + nameA + ...
                " comparisons are tested against subject-to-subject variation in the " + nameA + ...
                " effect. A Subject x " + nameB + " term cannot be separated from residual error and was omitted.";
        elseif nestedAinB
            reTerm = '(1|Subject) + (1|SubjB)';
            notes(end+1) = "Nested layout: each level of " + nameA + " occurs in only one level of " + nameB + ...
                ". Random effects: Subject and Subject x " + nameB + ", so " + nameB + ...
                " comparisons are tested against subject-to-subject variation in the " + nameB + ...
                " effect. A Subject x " + nameA + " term cannot be separated from residual error and was omitted.";
        else
            reTerm = '(1|Subject) + (1|SubjA) + (1|SubjB)';
        end
    end
end
if twoWay, fixedPart = 'Y ~ 1 + FA*FB'; else, fixedPart = 'Y ~ 1 + FA'; end
dfMethod = 'residual';
if isRM, dfMethod = 'satterthwaite'; end

mainMethod = '';
if fullDesign
    try
        [M, reTerm, notes] = fitWithFallback(T, fixedPart, reTerm, isRM, kind, ...
            {'DummyVarCoding', 'effects'}, '', notes);
    catch ME
        errordlg(ME.message, 'Model fit failed'); return;
    end
    mainMethod = M.method;
end

% Cell-means version of the same model (used for EMMs and post hoc tests)
Tc = T;  cellCol = zeros(nA, nB);  cNames = {};  cLabels = strings(0, 1);
for i = 1:nA
    for j = 1:nB
        if cellN(i, j) > 0
            if twoWay
                idx = T.FA == aLv{i} & T.FB == bLv{j};
                lbl = sprintf('%s %s | %s %s', nameA, aLv{i}, nameB, bLv{j});
            else
                idx = T.FA == aLv{i};
                lbl = sprintf('%s %s', nameA, aLv{i});
            end
            c = numel(cNames) + 1;
            cNames{c} = sprintf('Cell%d', c); %#ok<AGROW>
            Tc.(cNames{c}) = double(idx);
            cellCol(i, j) = c;
            cLabels(c, 1) = lbl;
        end
    end
end
cellFixed = ['Y ~ -1 + ' strjoin(cNames, ' + ')];
try
    [Mc, reTerm, notes] = fitWithFallback(Tc, cellFixed, reTerm, isRM, kind, {}, mainMethod, notes);
catch ME
    errordlg(['Model could not be fitted: ' ME.message], 'Model fit failed'); return;
end
if ~fullDesign
    M = Mc;
    fixedPart = ['Y ~ -1 + [one mean per ' nameA ' x ' nameB ' combination with data]'];
end
if isRM, fMain = [fixedPart ' + ' reTerm]; else, fMain = fixedPart; end
% readable formula for the output (internal names -> your factor names)
fMainShow = strrep(fMain, 'SubjA', ['Subject:' nameA]);
fMainShow = strrep(fMainShow, 'SubjB', ['Subject:' nameB]);
fMainShow = strrep(strrep(fMainShow, 'FA', nameA), 'FB', nameB);
fMainShow = strrep(fMainShow, 'Y ~', [respName ' ~']);

if isRM
    try
        fixedEffects(Mc.mdl, 'DFMethod', 'satterthwaite');
    catch
        dfMethod = 'residual';
        notes(end+1) = "Satterthwaite degrees of freedom could not be computed; residual df were used.";
    end
end

[bAll, bNames, Vall] = coefInfo(Mc);
pos = cellfun(@(s) find(strcmp(bNames, s), 1), cNames);
ctx = struct('beta', bAll(pos), 'V', Vall(pos, pos), 'M', Mc, 'pos', pos, ...
    'nAll', numel(bAll), 'dfMethod', dfMethod);

fprintf('\n===== %s =====\nFormula: %s\n', typeLabel, fMainShow);
disp(M.mdl);

%% ------------------------------------------------------------------------
%  6. Estimated marginal means and post hoc comparisons
%  ------------------------------------------------------------------------
emmSheets = {};  phSheets = {};
nC = numel(cNames);  E = eye(nC);
aLab = string(nameA) + " " + string(aLv);

LA = zeros(nA, nC);  noteA = strings(nA, 1);
for i = 1:nA
    cols = cellCol(i, cellCol(i, :) > 0);
    LA(i, cols) = 1 / numel(cols);
    if twoWay && numel(cols) < nB
        noteA(i) = sprintf('Averaged over the %d of %d levels of %s with data', numel(cols), nB, nameB);
    end
end
emmSheets(end+1, :) = {['EMM_' nameA], makeEMM(LA, nameA, aLab, noteA, ctx)};

if ~twoWay
    [L, lab, grp] = pairwiseRows(LA, aLab, "All levels of " + nameA);
    phSheets(end+1, :) = {['PH_' nameA], runFamily(L, lab, grp, ctx, "Pairwise " + nameA)};
    Lint = zeros(0, nC);  LB = zeros(0, nC);
else
    bLab = string(nameB) + " " + string(bLv);
    LB = zeros(nB, nC);  noteB = strings(nB, 1);
    for j = 1:nB
        cols = cellCol(cellCol(:, j) > 0, j)';
        LB(j, cols) = 1 / numel(cols);
        if numel(cols) < nA
            noteB(j) = sprintf('Averaged over the %d of %d levels of %s with data', numel(cols), nA, nameA);
        end
    end
    emmSheets(end+1, :) = {['EMM_' nameB], makeEMM(LB, nameB, bLab, noteB, ctx)};
    emmSheets(end+1, :) = {['EMM_' nameA '_x_' nameB], makeEMM(E, [nameA ' x ' nameB], cLabels, strings(nC, 1), ctx)};

    % main effects
    [L, lab, grp] = pairwiseRows(LA, aLab, "Averaged over " + nameB);
    phSheets(end+1, :) = {['PH_' nameA '_main'], runFamily(L, lab, grp, ctx, nameA + " (main effect)")};
    [L, lab, grp] = pairwiseRows(LB, bLab, "Averaged over " + nameA);
    phSheets(end+1, :) = {['PH_' nameB '_main'], runFamily(L, lab, grp, ctx, nameB + " (main effect)")};

    % A within each level of B
    L = zeros(0, nC);  lab = strings(0, 1);  grp = strings(0, 1);
    for j = 1:nB
        pres = find(cellCol(:, j) > 0);
        [l, b, g] = pairwiseRows(E(cellCol(pres, j), :), aLab(pres), "Within " + bLab(j));
        L = [L; l]; lab = [lab; b]; grp = [grp; g]; %#ok<AGROW>
    end
    phSheets(end+1, :) = {['PH_' nameA '_within_' nameB], runFamily(L, lab, grp, ctx, nameA + " within " + nameB + " (simple effects)")};

    % B within each level of A
    L = zeros(0, nC);  lab = strings(0, 1);  grp = strings(0, 1);
    for i = 1:nA
        pres = find(cellCol(i, :) > 0);
        [l, b, g] = pairwiseRows(E(cellCol(i, pres), :), bLab(pres), "Within " + aLab(i));
        L = [L; l]; lab = [lab; b]; grp = [grp; g]; %#ok<AGROW>
    end
    phSheets(end+1, :) = {['PH_' nameB '_within_' nameA], runFamily(L, lab, grp, ctx, nameB + " within " + nameA + " (simple effects)")};

    % interaction contrasts: (a1 - a2 at b_j) - (a1 - a2 at b_k)
    Lint = zeros(0, nC);  lab = strings(0, 1);
    tp = nchoosek(1:nA, 2);  dp = nchoosek(1:nB, 2);
    for p1 = 1:size(tp, 1)
        for p2 = 1:size(dp, 1)
            i1 = tp(p1, 1); i2 = tp(p1, 2); j1 = dp(p2, 1); j2 = dp(p2, 2);
            cc = [cellCol(i1, j1), cellCol(i2, j1), cellCol(i1, j2), cellCol(i2, j2)];
            if all(cc > 0)
                Lint(end+1, :) = E(cc(1), :) - E(cc(2), :) - E(cc(3), :) + E(cc(4), :); %#ok<AGROW>
                lab(end+1, 1) = sprintf('(%s - %s) at %s  vs  at %s', ...
                    aLab(i1), aLab(i2), bLab(j1), bLab(j2)); %#ok<AGROW>
            end
        end
    end
    grp = repmat("All interaction contrasts", numel(lab), 1);
    phSheets(end+1, :) = {'PH_Interaction_contrasts', runFamily(Lint, lab, grp, ctx, nameA + " x " + nameB + " interaction contrasts", false)};

    % every cell vs every other cell
    [L, lab, grp] = pairwiseRows(E, cLabels, "All " + nameA + " x " + nameB + " cells");
    phSheets(end+1, :) = {'PH_All_Cells', runFamily(L, lab, grp, ctx, "All pairwise " + nameA + " x " + nameB + " cells")};
end

% --- Optional: comparisons against a reference (many-to-one, Dunnett) ---
refInfo = '';
refAns = questdlg(['Also compare conditions against a REFERENCE (e.g. each test day vs baseline)?' newline newline ...
    'You choose the reference condition(s) - several are averaged into one baseline - ' ...
    'and the conditions to compare with it. P-values are Dunnett-adjusted.'], ...
    'Reference comparisons', 'Yes', 'No', 'No');
if strcmp(refAns, 'Yes')
    [refIdx, ok1] = listdlg('PromptString', {'Select the REFERENCE condition(s).', ...
        'Selecting several averages them into one baseline.'}, ...
        'ListString', cellstr(cLabels), 'SelectionMode', 'multiple', ...
        'ListSize', [460 300], 'Name', 'Reference');
    if ok1 && ~isempty(refIdx)
        others = setdiff(1:nC, refIdx);
        if isempty(others)
            warndlg('Every condition was selected as reference; no reference comparisons made.', 'Reference');
        else
            [cmpSel, ok2] = listdlg('PromptString', 'Select the conditions to compare with the reference:', ...
                'ListString', cellstr(cLabels(others)), 'SelectionMode', 'multiple', ...
                'InitialValue', 1:numel(others), 'ListSize', [460 300], 'Name', 'Compare with reference');
            if ok2 && ~isempty(cmpSel)
                cmpIdx = others(cmpSel);
                Lref = mean(E(refIdx, :), 1);
                if numel(refIdx) == 1
                    refLab = cLabels(refIdx);
                else
                    refLab = "mean of (" + string(strjoin(cellstr(cLabels(refIdx)), ', ')) + ")";
                end
                L   = E(cmpIdx, :) - Lref;
                lab = cLabels(cmpIdx) + " - " + refLab;
                grp = repmat("vs " + refLab, numel(cmpIdx), 1);
                phSheets(end+1, :) = {'PH_vs_Reference', ...
                    runDunnett(L, lab, grp, ctx, "Comparisons with reference (Dunnett)")};
                refInfo = char("Reference = " + refLab + "; compared: " + ...
                    string(strjoin(cellstr(cLabels(cmpIdx)), ', ')));
            end
        end
    end
end

%% ------------------------------------------------------------------------
%  7. ANOVA table and model summaries
%  ------------------------------------------------------------------------
if fullDesign
    aovTbl = termAnova(M, dfMethod, nameA, nameB);
else
    % Nested layouts: add a pooled test of the inner factor within the outer one
    extra = cell(0, 3);
    if nestedBinA && ~nestedAinB
        Hw = zeros(0, nC);
        for i = 1:nA
            cc = cellCol(i, cellCol(i, :) > 0);
            Hw = [Hw; E(cc(2:end), :) - E(cc(1), :)]; %#ok<AGROW>
        end
        extra(end+1, :) = {string(nameB) + " within " + string(nameA), Hw, ...
            "Differences among levels of " + nameB + " inside each level of " + nameA + ", pooled (nested design)."};
    elseif nestedAinB && ~nestedBinA
        Hw = zeros(0, nC);
        for j = 1:nB
            cc = cellCol(cellCol(:, j) > 0, j)';
            Hw = [Hw; E(cc(2:end), :) - E(cc(1), :)]; %#ok<AGROW>
        end
        extra(end+1, :) = {string(nameA) + " within " + string(nameB), Hw, ...
            "Differences among levels of " + nameA + " inside each level of " + nameB + ", pooled (nested design)."};
    end
    aovTbl = manualAnova(ctx, LA, LB, Lint, E, nameA, nameB, extra);
end
disp(aovTbl);

feTbl = fixedTable(M, dfMethod);
if ~fullDesign
    feNames = string(feTbl.Name);
    meaning = strings(height(feTbl), 1);
    for c = 1:nC
        meaning(feNames == cNames{c}) = "Mean of " + cLabels(c);
    end
    feTbl = addvars(feTbl, meaning, 'After', 1, 'NewVariableNames', 'Meaning');
end

mdl = M.mdl;
if isRM
    [psi, mse, cpStats] = covarianceParameters(mdl);
    reVar = table();
    for k = 1:numel(cpStats)
        reVar = [reVar; harmonizeCov(cpStats{k})]; %#ok<AGROW>
    end
    reVar.Variance = nan(height(reVar), 1);
    isStd = strcmpi(reVar.Type, "std");
    reVar.Variance(isStd) = reVar.Estimate(isStd) .^ 2;
    reVar.Group(reVar.Group == "SubjA") = "Subject x " + nameA;
    reVar.Group(reVar.Group == "SubjB") = "Subject x " + nameB;
    [~, ~, reStats] = randomEffects(mdl);
    reSubj = ds2table(reStats);
    varS = psi{1}(1, 1);
    if numel(psi) == 1
        icc = varS / (varS + mse);
    else
        icc = NaN;
    end
    fitted_ = fitted(mdl);
    resRaw  = residuals(mdl, 'ResidualType', 'Raw');
    resStd  = residuals(mdl, 'ResidualType', 'Standardized');
    modelDesc = 'Linear mixed model (fitlme, REML)';
else
    mse = mdl.MSE;
    fitted_ = mdl.Fitted;
    resRaw  = mdl.Residuals.Raw;
    resStd  = mdl.Residuals.Standardized;
    modelDesc = 'Linear model (fitlm), equivalent to classical ANOVA';
end

switch kind
    case 'oneway'
        equiv = 'Equivalent to classical one-way ANOVA (Type III).';
    case 'oneway_rm'
        equiv = ['Random intercept per subject. With complete data this gives the same results as one-way ' ...
            'RM ANOVA (sphericity assumed); subjects with missing values are kept.'];
    case 'twoway'
        equiv = 'Equivalent to classical two-way ANOVA with Type III sums of squares.';
    case 'twoway_rm'
        if nestedBinA || nestedAinB
            equiv = ['Nested repeated-measures layout (see Note rows). Random intercepts for Subject and for ' ...
                'Subject x the outer factor, so the outer factor is tested against subject-to-subject variation ' ...
                'in its effect; comparisons within one level of the outer factor are tested against residual error. ' ...
                'Subjects with missing values are kept.'];
        else
            equiv = ['Random intercepts for Subject, Subject x ' nameA ' and Subject x ' nameB ': the same error strata ' ...
                'as two-way RM ANOVA. With complete balanced data this gives the same results (sphericity assumed); ' ...
                'subjects with missing values are kept.'];
        end
    case 'mixed'
        equiv = ['Random intercept per subject (' nameA ' within, ' nameB ' between subjects). With complete data ' ...
            'this gives the same results as a two-way mixed (split-plot) ANOVA (sphericity assumed); subjects ' ...
            'with missing values are kept.'];
end

fit = {
    'Analysis',                     typeLabel
    'Model',                        modelDesc
    'Formula',                      fMainShow
    'Relation to classical ANOVA',  equiv
    'Fit method',                   M.method
    'Degrees of freedom method',    dfMethod
    'Observations used',            mdl.NumObservations
    'Log-likelihood',               mdl.LogLikelihood
    'AIC',                          mdl.ModelCriterion.AIC
    'BIC',                          mdl.ModelCriterion.BIC
    'Residual variance',            mse
    'Residual SD',                  sqrt(mse)
    'R-squared (ordinary)',         mdl.Rsquared.Ordinary
    'R-squared (adjusted)',         mdl.Rsquared.Adjusted
    'Fitting messages/warnings',    char(M.msg)};
if isRM
    fit(end+1, :) = {'Subjects', numel(categories(T.Subject))};
    fit(end+1, :) = {'Subject (random intercept) variance', varS};
    if ~isnan(icc)
        fit(end+1, :) = {'ICC (subject)', icc};
    else
        fit(end+1, :) = {'ICC', 'Several variance components: see Random_Effects_Variance'};
    end
end
if twoWay
    if fullDesign
        fit(end+1, :) = {'Design', sprintf('Complete: every %s x %s combination has data.', nameA, nameB)};
    else
        fit(end+1, :) = {'Design', sprintf(['Incomplete: %d of %d %s x %s combinations have no data ' ...
            '(see Design_Check). A cell-means model was fitted and ANOVA F-tests use only estimable contrasts.'], ...
            nnz(cellN == 0), numel(cellN), nameA, nameB)};
    end
end
for k = 1:numel(notes)
    fit(end+1, :) = {'Note', char(notes(k))}; %#ok<AGROW>
end
fitTbl = infoTable(fit);

% Descriptive statistics of the data used
if twoWay
    [G, gA, gB] = findgroups(T.FA, T.FB);
    desc = table(gA, gB, 'VariableNames', {nameA, nameB});
else
    [G, gA] = findgroups(T.FA);
    desc = table(gA, 'VariableNames', {nameA});
end
desc.N_obs = splitapply(@numel, T.Y, G);
if isRM, desc.N_subjects = splitapply(@(s) numel(unique(s)), T.Subject, G); end
desc.Mean   = splitapply(@mean, T.Y, G);
desc.SD     = splitapply(@std, T.Y, G);
desc.SEM    = desc.SD ./ sqrt(desc.N_obs);
desc.Median = splitapply(@median, T.Y, G);

% Data used, with fitted values and residuals
dataUsed = table(T.ExcelRow, T.Y, T.FA, 'VariableNames', {'ExcelRow', 'Value', nameA});
if twoWay, dataUsed.(nameB) = T.FB; end
if isRM,   dataUsed.Subject = T.Subject; end
dataUsed.Fitted                = fitted_;
dataUsed.Residual_raw          = resRaw;
dataUsed.Residual_standardized = resStd;

colInfo = {'Response column', respVar; ['Factor "' nameA '" column'], aVar};
if twoWay, colInfo(end+1, :) = {['Factor "' nameB '" column'], bVar}; end
if isRM,   colInfo(end+1, :) = {'Subject ID column', subjVar}; end
readme = [{
    'Analysis',              typeLabel
    'Input file',            inFile
    'Sheet',                 sheetName
    'Response name',         respName}; colInfo; {
    'Analysis date',         char(datetime('now'))
    'Model',                 [modelDesc '. ' equiv]
    'Missing data',          'Rows with a missing value in any column used were removed (see Missing_Data and Rows_Removed). All other observations were kept.'
    'Fixed-effect coding',   'Effects (sum-to-zero) coding; the ANOVA sheet gives Type III F-tests.'
    'Degrees of freedom',    'Repeated-measures designs: Satterthwaite approximation. Independent designs: residual df.'
    'EMMs',                  'Estimated marginal means, in the units of your response. Marginal means average cells equally.'
    'Post hoc tests',        't-tests of linear contrasts of the fixed effects. Difference = first minus second level.'
    'p_Bonferroni',          'Bonferroni'
    'p_Holm_Bonferroni',     'Holm-Bonferroni step-down (recommended general-purpose FWER control)'
    'p_Sidak',               'Sidak'
    'p_Hochberg',            'Hochberg step-up'
    'p_FDR_BH',              'Benjamini-Hochberg false discovery rate'
    'p_FDR_BY',              'Benjamini-Yekutieli FDR (valid under any dependence)'
    'p_Tukey',               'Tukey (Tukey-Kramer for unequal SEs): exact studentized range p-value with q = |t| x sqrt(2), k = number of means compared in the group, df = the comparison''s df. Same method as Origin, SPSS, GraphPad and R. Empty for families that are not all-pairwise (interaction contrasts).'
    'p_SingleStep_maxT',     'Single-step max-t adjustment from the joint distribution of the contrasts (100,000 simulations). The random seed is fixed (rng(1)), so re-running the same analysis gives identical values; the Monte Carlo error is about 0.001. Given for every family; for interaction contrasts it is the Tukey-type adjustment to use.'
    'p_Scheffe',             'Scheffe (protects all possible contrasts in the family)'
    'Adjustment scope',      'Adjustments are applied separately within each Group of each post hoc sheet.'
    'Data_Checks',           'Checks of subject IDs (duplicates in the same condition; subjects in more than one group).'
    'Design_Check',          'Number of observations in every factor combination.'
    'Assumption check',      'Data_Used contains fitted values and standardized residuals; plot them to check approximate normality and constant variance.'}];
if twoWay
    readme(end+1, :) = {'PH_All_Cells', 'Every cell mean vs every other. Adjusted p-values account for all of these comparisons, so they are stricter than in the other sheets.'};
end
if ~isempty(refInfo)
    readme(end+1, :) = {'PH_vs_Reference', ['Each selected condition minus the reference (Difference > 0 means higher than the reference). ' refInfo]};
    readme(end+1, :) = {'p_Dunnett', ['Dunnett (many-to-one) single-step adjustment over the comparisons in PH_vs_Reference: exact multivariate t ' ...
        'probability using the correlations between the comparisons and the smallest df in the family (conservative). ' ...
        'CI95 columns are unadjusted 95% intervals.']};
end
readmeTbl = infoTable(readme);

if isempty(checks), checks = {'Subject checks', 'Not applicable (independent groups design)'}; end
checksTbl = cell2table(checks, 'VariableNames', {'Check', 'Result'});

%% ------------------------------------------------------------------------
%  8. Export to Excel
%  ------------------------------------------------------------------------
defOut = fullfile(fp, [baseName '_' kind '_LMM_results.xlsx']);
[of, op] = uiputfile('*.xlsx', 'Save results as', defOut);
if isequal(of, 0)
    outFile = defOut;
    fprintf('No output location chosen; saving to %s\n', outFile);
else
    outFile = fullfile(op, of);
end
if isfile(outFile)
    try
        delete(outFile);
    catch
        errordlg('Could not overwrite the output file. Close it in Excel and run again.', 'LMM'); return;
    end
end

sheets = {'README', readmeTbl; 'Model_Fit', fitTbl; 'ANOVA_TypeIII', aovTbl; ...
          'Fixed_Effects', feTbl};
if isRM
    sheets = [sheets; {'Random_Effects_Variance', reVar; 'Random_Effects_Subjects', reSubj}];
end
sheets = [sheets; emmSheets; phSheets; ...
          {'Descriptives', desc; 'Design_Check', designCheck; 'Data_Checks', checksTbl; ...
           'Missing_Data', missSummary; 'Rows_Removed', removedRows; 'Data_Used', dataUsed}];

usedNames = strings(0, 1);
for k = 1:size(sheets, 1)
    tbl = sheets{k, 2};
    if isempty(tbl) || height(tbl) == 0
        tbl = table("No comparisons available (too few levels or empty cells).", 'VariableNames', {'Note'});
    end
    nm = sheetSafe(sheets{k, 1});
    base = nm;  n = 1;
    while any(usedNames == string(nm))
        n = n + 1;
        suffix = sprintf('_%d', n);
        nm = [base(1:min(end, 31 - numel(suffix))) suffix];
    end
    usedNames(end+1) = string(nm); %#ok<AGROW>
    writetable(tbl, outFile, 'Sheet', nm);
end
fprintf('\nResults written to:\n  %s\n', outFile);
msgbox(sprintf('Analysis finished.\n\nResults saved to:\n%s', outFile), 'LMM');
end


%% ========================================================================
%  Local functions
%  ========================================================================

function tok = missingTokens()
tok = {'NA', 'N/A', 'na', 'n/a', 'NaN', 'nan', '.', '-', '--', 'NULL', 'null', 'missing', '#N/A'};
end

function v = pickColumn(vn, prompt)
[k, ok] = listdlg('PromptString', prompt, 'ListString', vn, 'SelectionMode', 'single', ...
    'ListSize', [380 300], 'Name', 'Select column');
if ok, v = vn{k}; else, v = ''; disp('Cancelled.'); end
end

function s = cleanName(s, default)
s = strtrim(char(s));
if isempty(s), s = char(default); end
end

function s = sheetSafe(s)
s = regexprep(char(s), '[:\\/\?\*\[\]]', '_');
s = s(1:min(31, end));
end

function y = toNumeric(x)
% Convert any column type to double; unparseable entries become NaN.
if isnumeric(x) || islogical(x)
    y = double(x);
else
    s = strtrim(string(x));
    s(ismember(lower(s), lower(missingTokens())) | s == "") = missing;
    y = str2double(s);
end
y = y(:);
end

function c = toCategorical(x)
% Convert any column type to categorical; empty / NA entries become <undefined>.
if iscategorical(x)
    c = x;
elseif isnumeric(x) || islogical(x)
    c = categorical(double(x));
else
    s = strtrim(string(x));
    s(ismember(lower(s), lower(missingTokens())) | s == "") = missing;
    c = categorical(s);
end
c = removecats(c(:));
end

function c = toFactor(x)
% Categorical factor; numeric-looking levels are ordered numerically (1,2,...,10).
c = toCategorical(x);
s = string(c);
ok = ~ismissing(s);
v = str2double(s);
if any(ok) && all(~isnan(v(ok)))
    c = removecats(categorical(v));
end
end

function M = fitModel(T, formula, isMixed, extra, preferred)
% Ordinary linear model (fitlm) or linear mixed model (fitlme).
M = struct('mdl', [], 'isMixed', isMixed, 'method', '', 'msg', "None");
if ~isMixed
    lastwarn('');
    M.mdl = fitlm(T, formula, extra{:});
    M.method = 'Ordinary least squares';
    w = lastwarn;
    if ~isempty(w), M.msg = "Warning: " + string(w); end
    return
end
methods = {'REML', 'ML'};
if ~isempty(preferred) && any(strcmp(methods, preferred))
    methods = [{preferred}, setdiff(methods, {preferred}, 'stable')];
end
msg = "";
for k = 1:numel(methods)
    lastwarn('');
    try
        M.mdl = fitlme(T, formula, 'FitMethod', methods{k}, 'CheckHessian', true, extra{:});
        M.method = methods{k};
        w = lastwarn;
        if ~isempty(w), msg = msg + "Warning (" + M.method + "): " + string(w) + " "; end
        if msg == "", msg = "None"; end
        M.msg = msg;
        return
    catch ME
        msg = msg + "Failed with " + methods{k} + ": " + string(ME.message) + " ";
    end
end
error('LMM:fitFailed', 'The model could not be fitted. %s', msg);
end

function [M, reTerm, notes] = fitWithFallback(T, fixedPart, reTerm, isRM, kind, extra, preferred, notes)
% Fit; for two-way RM designs fall back to a subject intercept only if the
% full error structure cannot be estimated.
f = fixedPart;
if isRM, f = [f ' + ' reTerm]; end
try
    M = fitModel(T, f, isRM, extra, preferred);
catch ME
    if strcmp(kind, 'twoway_rm') && ~strcmp(reTerm, '(1|Subject)')
        notes(end+1) = "The full repeated-measures error structure could not be fitted (" + ...
            string(ME.message) + "). The model was fitted with (1|Subject) only.";
        reTerm = '(1|Subject)';
        M = fitModel(T, [fixedPart ' + ' reTerm], isRM, extra, preferred);
    else
        rethrow(ME);
    end
end
end

function [b, names, V] = coefInfo(M)
if M.isMixed
    [b, nd] = fixedEffects(M.mdl);
    names = cellstr(string(nd.Name));
else
    b = M.mdl.Coefficients.Estimate;
    names = M.mdl.CoefficientNames(:);
end
V = M.mdl.CoefficientCovariance;
end

function [p, F, df1, df2] = waldTest(M, H, dfMethod)
c = zeros(size(H, 1), 1);
if M.isMixed
    try
        [p, F, df1, df2] = coefTest(M.mdl, H, c, 'DFMethod', dfMethod);
    catch
        [p, F, df1, df2] = coefTest(M.mdl, H, c);
    end
else
    [p, F, df1] = coefTest(M.mdl, H, c);
    df2 = M.mdl.DFE;
end
end

function tbl = termAnova(M, dfMethod, nameA, nameB)
% Type III F-tests: all coefficients of each term jointly zero (effects coding).
[~, names] = coefInfo(M);
p = numel(names);
key = strings(p, 1);
for i = 1:p
    parts = split(string(names{i}), ':');
    v = strings(numel(parts), 1);
    for q = 1:numel(parts)
        tok = regexp(char(parts(q)), '^(FA|FB)_', 'tokens', 'once');
        if ~isempty(tok), v(q) = tok{1}; end
    end
    if all(v ~= ""), key(i) = strjoin(v, ':'); end
end
termKeys  = ["FA"; "FB"; "FA:FB"];
termNames = [string(nameA); string(nameB); string(nameA) + " x " + string(nameB)];
Term = strings(0, 1); [F, DF1, DF2, pv] = deal(zeros(0, 1));
for t = 1:numel(termKeys)
    idx = find(key == termKeys(t));
    if isempty(idx), continue; end
    H = zeros(numel(idx), p);
    H(sub2ind(size(H), (1:numel(idx))', idx)) = 1;
    [pp, ff, d1, d2] = waldTest(M, H, dfMethod);
    Term(end+1, 1) = termNames(t); F(end+1, 1) = ff; DF1(end+1, 1) = d1; %#ok<AGROW>
    DF2(end+1, 1) = d2; pv(end+1, 1) = pp; %#ok<AGROW>
end
tbl = table(Term, F, DF1, DF2, pv, 'VariableNames', {'Term', 'FStat', 'DF1', 'DF2', 'pValue'});
end

function tbl = manualAnova(ctx, LA, LB, Lint, E, nameA, nameB, extra)
% F-tests built from estimable contrasts of the cell means (incomplete designs).
if nargin < 8, extra = cell(0, 3); end
terms = [string(nameA); string(nameB); string(nameA) + " x " + string(nameB); "Any difference among all cell means"];
Hs = {LA(2:end, :) - LA(1, :), LB(2:end, :) - LB(1, :), Lint, E(2:end, :) - E(1, :)};
notes = ["Compares " + nameA + " marginal means; with empty cells each is averaged over the levels of " + nameB + " it has data for."; ...
         "Compares " + nameB + " marginal means; with empty cells each is averaged over the levels of " + nameA + " it has data for."; ...
         "Uses only interaction contrasts whose four cells all have data."; ...
         "Overall test that all cell means are equal."];
for k = 1:size(extra, 1)
    terms(end+1, 1) = extra{k, 1}; %#ok<AGROW>
    Hs{end+1}       = extra{k, 2}; %#ok<AGROW>
    notes(end+1, 1) = extra{k, 3}; %#ok<AGROW>
end
n = numel(terms);
[F, DF1, DF2, p] = deal(nan(n, 1));
for k = 1:n
    H = independentRows(Hs{k});
    if isempty(H)
        notes(k) = "Not estimable with this design. " + notes(k);
        continue
    end
    Hfull = zeros(size(H, 1), ctx.nAll);
    Hfull(:, ctx.pos) = H;
    [p(k), F(k), DF1(k), DF2(k)] = waldTest(ctx.M, Hfull, ctx.dfMethod);
end
tbl = table(terms, F, DF1, DF2, p, notes, ...
    'VariableNames', {'Term', 'FStat', 'DF1', 'DF2', 'pValue', 'Note'});
end

function H = independentRows(H)
% Keep a linearly independent subset of rows (removes redundant contrasts).
if isempty(H), return; end
H = H(any(abs(H) > 1e-12, 2), :);
if isempty(H), return; end
r = rank(H);
[~, ~, perm] = qr(H', 0);
H = H(sort(perm(1:r)), :);
end

function t = fixedTable(M, dfMethod)
if M.isMixed
    try
        [~, ~, st] = fixedEffects(M.mdl, 'DFMethod', dfMethod);
    catch
        [~, ~, st] = fixedEffects(M.mdl);
    end
    t = ds2table(st);
else
    t = M.mdl.Coefficients;
    t = addvars(t, string(t.Properties.RowNames), 'Before', 1, 'NewVariableNames', 'Name');
    t.Properties.RowNames = {};
    t.DF = repmat(M.mdl.DFE, height(t), 1);
end
end

function df = contrastDF(L, ctx)
% Degrees of freedom for each contrast row (Satterthwaite or residual).
m = size(L, 1);
df = repmat(ctx.M.mdl.DFE, m, 1);
if ~ctx.M.isMixed || ~strcmp(ctx.dfMethod, 'satterthwaite'), return; end
for r = 1:m
    H = zeros(1, ctx.nAll);
    H(ctx.pos) = L(r, :);
    try
        [~, ~, ~, df(r)] = coefTest(ctx.M.mdl, H, 0, 'DFMethod', 'satterthwaite');
    catch
        df(r) = ctx.M.mdl.DFE;
    end
end
end

function [L, lab, grp] = pairwiseRows(Lemm, labels, group)
% All pairwise differences between the rows of an EMM coefficient matrix.
labels = string(labels(:));
n = size(Lemm, 1);
if n < 2
    L = zeros(0, size(Lemm, 2)); lab = strings(0, 1); grp = strings(0, 1); return;
end
pr  = nchoosek(1:n, 2);
L   = Lemm(pr(:, 1), :) - Lemm(pr(:, 2), :);
lab = labels(pr(:, 1)) + " - " + labels(pr(:, 2));
grp = repmat(string(group), size(pr, 1), 1);
end

function tbl = makeEMM(L, factorName, levels, notes, ctx)
n    = size(L, 1);
est  = L * ctx.beta;
se   = sqrt(max(diag(L * ctx.V * L'), 0));
df   = contrastDF(L, ctx);
crit = tinv(0.975, df);
lo = est - crit .* se;  hi = est + crit .* se;
tbl = table(repmat(string(factorName), n, 1), string(levels(:)), est, se, df, lo, hi, ...
    string(notes(:)), 'VariableNames', {'Factor', 'Level', 'Mean', 'SE', 'df', ...
    'CI95_lower', 'CI95_upper', 'Note'});
end

function tbl = runFamily(L, labels, groups, ctx, familyName, isPairwise)
% t-tests of the contrasts in L with all p-value adjustments.
% isPairwise = true when each group holds all pairwise differences of k means
% (exact Tukey p-values are then computed); false for other contrast families.
if nargin < 6, isPairwise = true; end
tbl = table();
if isempty(L), return; end
m    = size(L, 1);
est  = L * ctx.beta;
C    = L * ctx.V * L';
se   = sqrt(max(diag(C), 0));
df   = contrastDF(L, ctx);
tt   = est ./ se;
p    = 2 * tcdf(-abs(tt), df);
crit = tinv(0.975, df);
lo = est - crit .* se;  hi = est + crit .* se;

[pBonf, pHolmBonf, pSidak, pHoch, pBH, pBY, pTuk, pMaxT, pSch, qTuk] = deal(nan(m, 1));
kTuk = nan(m, 1);
ug = unique(groups, 'stable');
for g = 1:numel(ug)
    ix = find(groups == ug(g) & ~isnan(p));
    if isempty(ix), continue; end
    pp = p(ix);
    pBonf(ix)  = padjust(pp, 'bonferroni');
    pHolmBonf(ix) = padjust(pp, 'holm-bonferroni');
    pSidak(ix) = padjust(pp, 'sidak');
    pHoch(ix)  = padjust(pp, 'hochberg');
    pBH(ix)    = padjust(pp, 'BH');
    pBY(ix)    = padjust(pp, 'BY');
    pMaxT(ix)  = singleStep(tt(ix), C(ix, ix), min(df(ix)));   % simulated, conservative common df
    % Exact Tukey / Tukey-Kramer: m = k(k-1)/2 pairwise comparisons of k means
    nk = numel(ix);
    kk = (1 + sqrt(1 + 8 * nk)) / 2;
    if isPairwise && abs(kk - round(kk)) < 1e-9
        kTuk(ix) = round(kk);
        qTuk(ix) = abs(tt(ix)) * sqrt(2);
        pTuk(ix) = ptukeyUpper(qTuk(ix), round(kk), df(ix));
    end
    r = max(rank(L(ix, :)), 1);
    pSch(ix)   = 1 - fcdf(tt(ix) .^ 2 ./ r, r, df(ix));
end

tbl = table(repmat(string(familyName), m, 1), groups(:), labels(:), est, se, df, tt, lo, hi, ...
    p, qTuk, kTuk, pTuk, pBonf, pHolmBonf, pSidak, pHoch, pBH, pBY, pMaxT, pSch, ...
    'VariableNames', {'Family', 'Group', 'Comparison', 'Difference', 'SE', 'df', 't', ...
    'CI95_lower', 'CI95_upper', 'p_unadjusted', 'q_Tukey', 'k_means', 'p_Tukey', ...
    'p_Bonferroni', 'p_Holm_Bonferroni', 'p_Sidak', 'p_Hochberg', 'p_FDR_BH', 'p_FDR_BY', ...
    'p_SingleStep_maxT', 'p_Scheffe'});
tbl.Significant_Holm_Bonferroni_0p05 = pHolmBonf < 0.05;
tbl.Significant_FDR_0p05  = pBH < 0.05;
end

function p = ptukeyUpper(q, k, df)
% Upper-tail probability P(Q >= q) of the studentized range distribution with
% k means and df error degrees of freedom (the Tukey / Tukey-Kramer p-value).
% Computed by numerical integration over the distribution of s = sqrt(chi2/df).
p = nan(size(q));
for i = 1:numel(q)
    if ~isfinite(q(i)), continue; end
    if q(i) <= 0, p(i) = 1; continue; end
    nu = df(min(i, numel(df)));
    if ~isfinite(nu) || nu > 1e6
        p(i) = rangeUpper(q(i), k);
    else
        f = @(s) tukeyIntegrand(s, q(i), k, nu);
        p(i) = integral(f, 0, Inf, 'AbsTol', 1e-14, 'RelTol', 1e-10);
    end
    p(i) = min(max(p(i), 0), 1);
end
end

function v = tukeyIntegrand(s, q, k, nu)
% density of s = sqrt(chi2_nu / nu) times P(range >= q*s)
logc = (nu / 2) * log(nu) - gammaln(nu / 2) - (nu / 2 - 1) * log(2);
dens = exp(logc + (nu - 1) .* log(s) - nu .* s .^ 2 ./ 2);
dens(~isfinite(dens)) = 0;
v = dens .* rangeUpper(q .* s, k);
end

function P = rangeUpper(w, k)
% P(range of k independent standard normals >= w), for each element of w.
% Uses a^(k-1) - b^(k-1) = c * sum(a^i b^(k-2-i)) so small tails stay accurate.
sz = size(w);  w = w(:)';
P = zeros(size(w));
P(w <= 0) = 1;
use = w > 0 & w < 40;                    % beyond 40 the probability is ~0
if any(use)
    wp = w(use);
    h  = 0.005;
    z  = (-8:h:(8 + max(wp)))';
    a  = normcdf(z);
    c  = normcdf(z - wp);                % numel(z) x numel(wp)
    b  = a - c;
    S  = zeros(size(c));
    for i = 0:k-2
        S = S + a .^ i .* b .^ (k - 2 - i);
    end
    P(use) = k * trapz(z, normpdf(z) .* c .* S, 1);
end
P = reshape(P, sz);
end

function tbl = runDunnett(L, labels, groups, ctx, familyName)
% Many-to-one comparisons: all runFamily columns plus a Dunnett-adjusted p.
tbl = runFamily(L, labels, groups, ctx, familyName, false);
if isempty(tbl) || height(tbl) == 0, return; end
k  = size(L, 1);
C  = L * ctx.V * L';
se = sqrt(max(diag(C), 0));
R  = C ./ (se * se');  R = (R + R') / 2;
nu = min(tbl.df);
pD = nan(k, 1);
opts = statset('TolFun', 1e-5);
for i = 1:k
    a = abs(tbl.t(i));
    if ~isfinite(a), continue; end
    if k == 1
        pD(i) = 2 * tcdf(-a, nu);
    elseif k <= 25
        try
            pD(i) = 1 - mvtcdf(-a * ones(1, k), a * ones(1, k), R, nu, opts);
        catch
            pD(i) = NaN;
        end
    end
end
miss = isnan(pD) & isfinite(tbl.t);
if any(miss)                     % fall back to simulation (large families / numerical failure)
    pSim = singleStep(tbl.t, C, nu);
    pD(miss) = pSim(miss);
end
pD = min(max(pD, 0), 1);
tbl = addvars(tbl, pD, 'After', 'p_unadjusted', 'NewVariableNames', 'p_Dunnett');
end

function padj = padjust(p, method)
p = p(:);  n = numel(p);  padj = nan(n, 1);
switch lower(method)
    case 'bonferroni'
        padj = min(p * n, 1);
    case 'sidak'
        padj = 1 - (1 - p) .^ n;
    case 'holm-bonferroni'
        [ps, ix] = sort(p, 'ascend');
        a = cummax(ps .* (n:-1:1)');
        padj(ix) = min(a, 1);
    case 'hochberg'
        [ps, ix] = sort(p, 'descend');
        a = cummin(ps .* (1:n)');
        padj(ix) = min(a, 1);
    case 'bh'
        [ps, ix] = sort(p, 'descend');
        a = cummin(ps .* n ./ (n:-1:1)');
        padj(ix) = min(a, 1);
    case 'by'
        cn = sum(1 ./ (1:n));
        [ps, ix] = sort(p, 'descend');
        a = cummin(ps .* n * cn ./ (n:-1:1)');
        padj(ix) = min(a, 1);
end
end

function padj = singleStep(t, C, df)
% Single-step max-|t| adjusted p-values (Tukey-type), by simulation from the
% joint multivariate t distribution of the contrasts.
t = t(:);  m = numel(t);  padj = nan(m, 1);
d = sqrt(diag(C));
ok = d > 0 & isfinite(t);
if ~any(ok), return; end
t = t(ok);  C = C(ok, ok);  d = d(ok);  k = numel(t);
R = C ./ (d * d');  R = (R + R') / 2;
[Q, Lam] = eig(R);
A = Q * diag(sqrt(max(diag(Lam), 0)));
nSim = 100000;  chunk = max(1000, floor(2e6 / k));
cnt = zeros(k, 1);  done = 0;  at = abs(t)';
while done < nSim
    n = min(chunk, nSim - done);
    Z = randn(n, k) * A';
    if isfinite(df), Z = Z ./ sqrt(chi2rnd(df, n, 1) / df); end
    mx = max(abs(Z), [], 2);
    cnt = cnt + sum(mx >= at, 1)';
    done = done + n;
end
padj(ok) = (cnt + 1) / (nSim + 1);
end

function t = ds2table(ds)
% Convert the dataset-style outputs of fitlme to a table.
if istable(ds), t = ds; return; end
try
    t = dataset2table(ds); return;
catch
end
try
    names = ds.Properties.VarNames;
catch
    names = get(ds, 'VarNames');
end
t = table();
for k = 1:numel(names)
    v = ds.(names{k});
    if ischar(v), v = cellstr(v); end
    t.(names{k}) = v;
end
end

function t = harmonizeCov(ds)
% Put the covariance-parameter tables for each grouping level into one layout.
t0 = ds2table(ds);  n = height(t0);  vn = t0.Properties.VariableNames;
t = table();
for f = {'Group', 'Name1', 'Name2', 'Type'}
    f1 = f{1};
    if ismember(f1, vn)
        t.(f1) = string(t0.(f1));
    elseif strcmp(f1, 'Name1') && ismember('Name', vn)
        t.(f1) = string(t0.Name);
    else
        t.(f1) = repmat("", n, 1);
    end
end
for f = {'Estimate', 'Lower', 'Upper'}
    if ismember(f{1}, vn), t.(f{1}) = double(t0.(f{1})); else, t.(f{1}) = nan(n, 1); end
end
end

function t = infoTable(c)
vals = cell(size(c, 1), 1);
for k = 1:size(c, 1)
    v = c{k, 2};
    if isnumeric(v) || islogical(v)
        vals{k} = num2str(v, 8);
    else
        vals{k} = char(string(v));
    end
end
t = table(string(c(:, 1)), string(vals), 'VariableNames', {'Item', 'Value'});
end
