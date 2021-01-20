using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DU_Industry_Price_Calculator
{
    public partial class PriceInputForm : Form
    {
        // Makes a form to input prices for the given recipes
        // Modified recipes can be extracted from these props

        private List<TextBox> inputs = new List<TextBox>();
        public IEnumerable<Recipe> final_recipes { get; private set; }

        public PriceInputForm(IEnumerable<Recipe> recipes)
        {
            this.final_recipes = recipes;
            InitializeComponent();
            this.AutoSize = true;
            // Make a layout container that will arrange them vertically
            var flowPanel = new FlowLayoutPanel();
            flowPanel.AutoSize = true;
            flowPanel.WrapContents = false;
            flowPanel.FlowDirection = FlowDirection.TopDown;
            foreach(var recipe in recipes)
            {
                // Make a layout container to arrange these left to right
                var recipePanel = new FlowLayoutPanel();
                recipePanel.AutoSize = true;
                recipePanel.WrapContents = false;
                recipePanel.FlowDirection = FlowDirection.LeftToRight;

                var recipeLabel = new Label();
                recipeLabel.Text = recipe.Name;
                recipePanel.Controls.Add(recipeLabel);

                var recipeInput = new TextBox();
                if (recipe.Price > 0)
                    recipeInput.Text = recipe.Price.ToString(); // Let prices be updated if necessary
                recipeInput.Tag = recipe;
                recipePanel.Controls.Add(recipeInput);
                inputs.Add(recipeInput);

                flowPanel.Controls.Add(recipePanel);
            }
            var buttonPanel = new FlowLayoutPanel();
            buttonPanel.AutoSize = true;
            buttonPanel.WrapContents = false;
            buttonPanel.FlowDirection = FlowDirection.LeftToRight;
            // Add a submit button
            var submitButton = new Button();
            submitButton.Text = "Submit";
            submitButton.Click += SubmitButton_Click;
            this.AcceptButton = submitButton;
            buttonPanel.Controls.Add(submitButton);

            var cancelButton = new Button();
            cancelButton.Text = "Cancel";
            this.CancelButton = cancelButton;
            buttonPanel.Controls.Add(cancelButton);

            flowPanel.Controls.Add(buttonPanel);

            this.Controls.Add(flowPanel);

            this.Refresh();
        }

        private void SubmitButton_Click(object sender, EventArgs e)
        {
            // We have a collection of TextBoxes; each one has a Tag that is its Recipe
            // We can construct our own list to replace the original
            List<Recipe> newRecipes = new List<Recipe>();
            foreach(var textbox in inputs)
            {
                var recipe = (Recipe)textbox.Tag;
                if (float.TryParse(textbox.Text, out float value))
                {
                    recipe.Price = value;
                }                
                newRecipes.Add(recipe); // Do this even if they didn't give us a price

            }

            final_recipes = newRecipes;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
