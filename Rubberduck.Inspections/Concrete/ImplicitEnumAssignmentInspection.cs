using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ImplicitEnumAssignmentInspection : ParseTreeInspectionBase
    {
        public ImplicitEnumAssignmentInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint) { }

        public override CodeInspectionType InspectionType { get; } = CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public override IInspectionListener Listener { get; } = new ImplicitEnumAssignmentInspectionLister();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts.Select(context => new QualifiedContextInspectionResult(this,
                                                                string.Format(InspectionsUI.ImplicitEnumAssignmentInspectionResultFormat, context.Context.start.Text),
                                                                context));
        }
    }

    public class ImplicitEnumAssignmentInspectionLister : VBAParserBaseListener, IInspectionListener
    {
        private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();

        public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

        public QualifiedModuleName CurrentModuleName { get; set; }

        public void ClearContexts()
        {
            _contexts.Clear();
        }

        public override void ExitEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            foreach (var enumeratorContext in context.enumerationStmt_Constant())
            {
                if (enumeratorContext.expression() == null)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, enumeratorContext));
                }
            }
        }
    }
}
