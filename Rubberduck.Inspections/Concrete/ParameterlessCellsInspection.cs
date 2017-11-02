using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
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
    public sealed class ParameterlessCellsInspection : ParseTreeInspectionBase
    {
        public ParameterlessCellsInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint) { }

        public override CodeInspectionType InspectionType { get; } = CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public override IInspectionListener Listener { get; } = new ParameterlessCellsInspectionLister();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var cellsReferences = State.DeclarationFinder.MatchName("Cells")
                .Where(member => member.AsTypeName == "Range"
                                 && member.References.Any()
                                 && member.ParentDeclaration.AsTypeName == "Range")
                .SelectMany(declaration => declaration.References);

            var context1 = Listener.Contexts.Where(context => !IsIgnoringInspectionResultFor(context.ModuleName, context.Context.Start.Line));

            return null;
        }
    }

    public class ParameterlessCellsInspectionLister : VBAParserBaseListener, IInspectionListener
    {
        private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();

        public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

        public QualifiedModuleName CurrentModuleName { get; set; }

        public void ClearContexts()
        {
            _contexts.Clear();
        }

        public override void ExitArgList(VBAParser.ArgListContext context)
        {

        }
    }
}
